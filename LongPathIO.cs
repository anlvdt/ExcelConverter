// Native Unicode filesystem calls keep Windows PowerShell 5.1 independent of MAX_PATH.
using System;
using System.IO;
using System.Collections.Generic;
using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using Microsoft.Win32.SafeHandles;
namespace ExcelConverter {
    public sealed class FileEntry {
        public string FullName, Name, Directory, BaseName, Extension;
        public bool IsDirectory, IsLink;
        public FileEntry(string path, bool directory, bool link) {
            FullName = path; Name = LongIO.Name(path); Directory = LongIO.Parent(path);
            Extension = Path.GetExtension(Name); BaseName = Path.GetFileNameWithoutExtension(Name);
            IsDirectory = directory; IsLink = link;
        }
    }
    public static class LongIO {
        static bool Windows { get { return Environment.OSVersion.Platform == PlatformID.Win32NT; } }
        public static string FullPath(string path) {
            if (String.IsNullOrWhiteSpace(path)) throw new ArgumentException("Empty path");
            if (Windows && path.StartsWith(@"\\?\UNC\", StringComparison.OrdinalIgnoreCase)) path = @"\\" + path.Substring(8);
            else if (Windows && path.StartsWith(@"\\?\")) path = path.Substring(4);
            if (!Windows) return Path.GetFullPath(path);
            if (path.StartsWith(@"\\.\")) throw new ArgumentException("Device paths are unsupported");
            path = path.Replace('/', '\\');
            string root; string tail;
            if (path.StartsWith(@"\\")) {
                string[] parts = path.Substring(2).Split(new char[] { '\\' }, StringSplitOptions.RemoveEmptyEntries);
                if (parts.Length < 2) throw new ArgumentException("UNC path must include a server and share");
                root = @"\\" + parts[0] + "\\" + parts[1] + "\\";
                tail = String.Join("\\", parts, 2, parts.Length - 2);
            } else if (path.Length >= 3 && Char.IsLetter(path[0]) && path[1] == ':' && path[2] == '\\') {
                root = path.Substring(0, 3); tail = path.Substring(3);
            } else {
                if (path.Contains(":")) throw new ArgumentException("Use a fully qualified drive path");
                string current = Environment.CurrentDirectory;
                if (path.StartsWith("\\")) return FullPath(Path.GetPathRoot(current) + path.TrimStart('\\'));
                return FullPath(current.TrimEnd('\\') + "\\" + path);
            }
            var segments = new List<string>();
            foreach (string part in tail.Split(new char[] { '\\' }, StringSplitOptions.RemoveEmptyEntries)) {
                if (part == ".") continue;
                if (part == "..") { if (segments.Count > 0) segments.RemoveAt(segments.Count - 1); continue; }
                if (part.Contains(":")) throw new ArgumentException("Alternate data streams are unsupported");
                segments.Add(part);
            }
            return root + String.Join("\\", segments.ToArray());
        }

        public static string Name(string path) {
            if (!Windows) return Path.GetFileName(path);
            return path.Substring(path.LastIndexOf('\\') + 1);
        }
        public static string Parent(string path) {
            if (!Windows) return Path.GetDirectoryName(path);
            int index = path.TrimEnd('\\').LastIndexOf('\\');
            if (index < 0) return null;
            return index == 2 && path[1] == ':' ? path.Substring(0, 3) : path.Substring(0, index);
        }

        static string Native(string p) {
            p = FullPath(p);
            if (!Windows) return p;
            return p.StartsWith(@"\\") ? @"\\?\UNC\" + p.Substring(2) : @"\\?\" + p;
        }
        static void Check(bool ok, string operation) { if (!ok) throw new Win32Exception(Marshal.GetLastWin32Error(), operation); }
        public static bool Exists(string path) {
            if (!Windows) return File.Exists(path) || System.IO.Directory.Exists(path);
            uint a = GetFileAttributesW(Native(path));
            if (a != 0xffffffff) return true;
            int code = Marshal.GetLastWin32Error();
            if (code == 2 || code == 3) return false;
            throw new Win32Exception(code, "Cannot inspect path: " + path);
        }
        public static bool IsDirectory(string path) {
            if (!Windows) return System.IO.Directory.Exists(path);
            uint a = GetFileAttributesW(Native(path));
            if (a == 0xffffffff) throw new Win32Exception(Marshal.GetLastWin32Error(), "Cannot inspect: " + path);
            return (a & 16) != 0;
        }
        public static FileEntry[] List(string directory) {
            var entries = new List<FileEntry>();
            if (!Windows) {
                foreach (string p in System.IO.Directory.GetFileSystemEntries(directory)) {
                    FileAttributes a = File.GetAttributes(p);
                    entries.Add(new FileEntry(p, (a & FileAttributes.Directory) != 0, (a & FileAttributes.ReparsePoint) != 0));
                }
                return entries.ToArray();
            }
            FindData data;
            IntPtr handle = FindFirstFileW(Native(directory).TrimEnd('\\') + @"\*", out data);
            if (handle == new IntPtr(-1)) {
                int code = Marshal.GetLastWin32Error();
                if (code == 2) return entries.ToArray();
                throw new Win32Exception(code, "Cannot enumerate: " + directory);
            }
            try {
                do {
                    if (data.Name == "." || data.Name == "..") continue;
                    // Cloud placeholders are reparse points too: skip only symlinks/junctions.
                    bool link = data.Tag == 0xa000000c || data.Tag == 0xa0000003;
                    entries.Add(new FileEntry(Path.Combine(directory, data.Name), (data.Attributes & 16) != 0, link));
                } while (FindNextFileW(handle, out data));
                int code = Marshal.GetLastWin32Error();
                if (code != 18) throw new Win32Exception(code, "Enumeration interrupted: " + directory);
            } finally { FindClose(handle); }
            return entries.ToArray();
        }
        public static void CreateDirectory(string path) {
            if (!Windows) { System.IO.Directory.CreateDirectory(path); return; }
            if (Exists(path)) { if (!IsDirectory(path)) throw new IOException("Not a directory: " + path); return; }
            string parent = Parent(path);
            if (!String.IsNullOrEmpty(parent) && !Exists(parent)) CreateDirectory(parent);
            if (!CreateDirectoryW(Native(path), IntPtr.Zero)) {
                int code = Marshal.GetLastWin32Error();
                if (code != 183 || !IsDirectory(path)) throw new Win32Exception(code, "Cannot create: " + path);
            }
        }
        public static void DeleteTree(string path) {
            foreach (FileEntry entry in List(path)) {
                if (entry.IsDirectory && !entry.IsLink) DeleteTree(entry.FullName);
                else if (entry.IsDirectory) RemoveDirectory(entry.FullName);
                else Delete(entry.FullName);
            }
            RemoveDirectory(path);
        }
        static void RemoveDirectory(string path) {
            if (Windows) Check(RemoveDirectoryW(Native(path)), "Cannot remove directory: " + path);
            else System.IO.Directory.Delete(path);
        }
        public static void Copy(string source, string destination) {
            if (Windows) Check(CopyFileW(Native(source), Native(destination), true), "Cannot copy: " + source + " -> " + destination);
            else File.Copy(source, destination, false);
        }
        public static void Move(string source, string destination) {
            if (Windows) Check(MoveFileW(Native(source), Native(destination)), "Cannot move: " + source + " -> " + destination);
            else File.Move(source, destination);
        }
        public static void Delete(string path) {
            if (Windows) Check(DeleteFileW(Native(path)), "Cannot delete: " + path);
            else File.Delete(path);
        }
        public static Stream OpenRead(string path) {
            if (!Windows) return new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);
            SafeFileHandle h = CreateFileW(Native(path), 0x80000000, 1, IntPtr.Zero, 3, 0x08000000, IntPtr.Zero);
            if (h.IsInvalid) { int code = Marshal.GetLastWin32Error(); h.Dispose(); throw new Win32Exception(code, "Cannot read: " + path); }
            try { return new FileStream(h, FileAccess.Read); } catch { h.Dispose(); throw; }
        }
        public static string Hash(string path) {
            using (Stream stream = OpenRead(path)) using (SHA256 sha = SHA256.Create())
                return BitConverter.ToString(sha.ComputeHash(stream)).Replace("-", "");
        }
        [StructLayout(LayoutKind.Sequential, CharSet=CharSet.Unicode)]
        struct FindData {
            public uint Attributes; public System.Runtime.InteropServices.ComTypes.FILETIME Creation, Access, Write;
            public uint SizeHigh, SizeLow, Tag, Reserved;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=260)] public string Name;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst=14)] public string AlternateName;
        }
        [DllImport("kernel32.dll", CharSet=CharSet.Unicode, SetLastError=true)] static extern uint GetFileAttributesW(string p);
        [DllImport("kernel32.dll", CharSet=CharSet.Unicode, SetLastError=true)] static extern IntPtr FindFirstFileW(string p, out FindData data);
        [DllImport("kernel32.dll", CharSet=CharSet.Unicode, SetLastError=true)] static extern bool FindNextFileW(IntPtr h, out FindData data);
        [DllImport("kernel32.dll")] static extern bool FindClose(IntPtr h);
        [DllImport("kernel32.dll", CharSet=CharSet.Unicode, SetLastError=true)] static extern bool CopyFileW(string s, string d, bool failIfExists);
        [DllImport("kernel32.dll", CharSet=CharSet.Unicode, SetLastError=true)] static extern bool MoveFileW(string s, string d);
        [DllImport("kernel32.dll", CharSet=CharSet.Unicode, SetLastError=true)] static extern bool RemoveDirectoryW(string p);
        [DllImport("kernel32.dll", CharSet=CharSet.Unicode, SetLastError=true)] static extern bool DeleteFileW(string p);
        [DllImport("kernel32.dll", CharSet=CharSet.Unicode, SetLastError=true)] static extern bool CreateDirectoryW(string p, IntPtr security);
        [DllImport("kernel32.dll", CharSet=CharSet.Unicode, SetLastError=true)] static extern SafeFileHandle CreateFileW(string p, uint access, uint share, IntPtr security, uint creation, uint flags, IntPtr template);
    }
}
