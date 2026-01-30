# XLSStart Virus Cleaner - User Guide
# Hướng dẫn sử dụng Công cụ Quét Virus XLSTART

Author: Le An (Vietnam IT)  
Version: 1.0.0  
Date: January 2026

---

## What is XLSStart Cleaner? | XLSStart Cleaner là gì?

**English:**
XLSStart Cleaner is a standalone security tool designed to detect and remove X97M/Laroux virus from Excel XLSTART folders. Unlike the main Excel Converter which converts files, this tool focuses solely on virus detection and removal without modifying your Excel files.

**Tiếng Việt:**
XLSStart Cleaner là công cụ bảo mật độc lập được thiết kế để phát hiện và xóa virus X97M/Laroux từ thư mục XLSTART của Excel. Khác với Excel Converter chính chuyển đổi file, công cụ này chỉ tập trung vào việc phát hiện và xóa virus mà không sửa đổi file Excel của bạn.

---

## Why do you need this? | Tại sao bạn cần công cụ này?

**English:**
The X97M/Laroux virus infects Excel by placing malicious files in XLSTART folders. These files automatically load every time Excel starts, spreading the infection to all workbooks you open. This tool:
- Scans 13+ possible XLSTART locations across all Office versions (2007-365)
- Detects 27+ known virus file variants
- Performs deep VBA code analysis to find hidden infections
- Removes infected files safely

**Tiếng Việt:**
Virus X97M/Laroux lây nhiễm Excel bằng cách đặt các file độc hại vào thư mục XLSTART. Các file này tự động tải mỗi khi Excel khởi động, lan truyền nhiễm sang tất cả workbook bạn mở. Công cụ này:
- Quét 13+ vị trí XLSTART có thể có trên tất cả phiên bản Office (2007-365)
- Phát hiện 27+ biến thể virus đã biết
- Phân tích sâu VBA code để tìm nhiễm ẩn
- Xóa file nhiễm virus một cách an toàn

---

## How to use | Cách sử dụng

### Method 1: From Excel Converter GUI | Phương pháp 1: Từ GUI Excel Converter

1. Open Excel Converter (run `ExcelConverter.cmd`)
2. Click the **🛡 Scan XLSTART** button in the header
3. XLSStart Cleaner will open in a new window

**Vietnamese:**
1. Mở Excel Converter (chạy `ExcelConverter.cmd`)
2. Nhấn nút **🛡 Scan XLSTART** ở phần header
3. XLSStart Cleaner sẽ mở trong cửa sổ mới

### Method 2: Run Standalone | Phương pháp 2: Chạy độc lập

**Option A: Using Batch File | Tùy chọn A: Dùng file Batch**
- Double-click `XLSStartCleaner.bat`
- Nhấp đúp vào `XLSStartCleaner.bat`

**Option B: Using PowerShell | Tùy chọn B: Dùng PowerShell**
```powershell
powershell.exe -ExecutionPolicy Bypass -File "XLSStartCleaner.ps1"
```

---

## User Interface | Giao diện người dùng

### Buttons | Các nút

**SCAN ONLY**
- **English:** Scans all XLSTART folders and reports infected files WITHOUT removing them
- **Vietnamese:** Quét tất cả thư mục XLSTART và báo cáo file nhiễm virus NHƯNG KHÔNG xóa chúng
- **Use when:** You want to check for infections first before taking action
- **Dùng khi:** Bạn muốn kiểm tra nhiễm virus trước khi thực hiện hành động

**SCAN & CLEAN**
- **English:** Scans all XLSTART folders and PERMANENTLY DELETES infected files
- **Vietnamese:** Quét tất cả thư mục XLSTART và XÓA VĨNH VIỄN file nhiễm virus
- **Use when:** You want to remove all detected viruses
- **Dùng khi:** Bạn muốn xóa tất cả virus đã phát hiện
- ⚠️ **WARNING:** This action cannot be undone! | Hành động này không thể hoàn tác!

**CLOSE**
- **English:** Exit the application
- **Vietnamese:** Thoát ứng dụng

---

## What gets scanned? | Những gì được quét?

### XLSTART Locations | Vị trí XLSTART

The tool scans these locations:

**User Profile:**
- `%APPDATA%\Microsoft\Excel\XLSTART`
- `%USERPROFILE%\AppData\Roaming\Microsoft\Excel\XLSTART`

**Office 365 / 2019 / 2021:**
- `C:\Program Files\Microsoft Office\root\Office16\XLSTART`
- `C:\Program Files (x86)\Microsoft Office\root\Office16\XLSTART`

**Office 2016:**
- `C:\Program Files\Microsoft Office\Office16\XLSTART`
- `C:\Program Files (x86)\Microsoft Office\Office16\XLSTART`

**Office 2013:**
- `C:\Program Files\Microsoft Office\Office15\XLSTART`
- `C:\Program Files (x86)\Microsoft Office\Office15\XLSTART`

**Office 2010:**
- `C:\Program Files\Microsoft Office\Office14\XLSTART`
- `C:\Program Files (x86)\Microsoft Office\Office14\XLSTART`

**Office 2007:**
- `C:\Program Files\Microsoft Office\Office12\XLSTART`
- `C:\Program Files (x86)\Microsoft Office\Office12\XLSTART`

**Legacy:**
- `C:\MSOFFICE\EXCEL\XLSTART`

**Special Files:**
- `C:\Windows\System\Xlscan.386` (VCX variant)

---

## Virus Detection Database | Cơ sở dữ liệu phát hiện virus

### Known Virus Filenames (27+) | Tên file virus đã biết (27+)

```
BINV.XLS, BOOK1.XLS, CAR.XLS, CURE.XLS, DIMON.XLS,
ECSYSTEM.XLS, KINSLAYER.XLS, NEGS.XLS, NOCAL.XLS,
PERSONAL.XLS, PLDT.XLS, PRIVAT.XLS, RESULTS.XLS,
SGV.XLS, SING.XLS, STARTUP.XLS, VERA.XLS, WINDOS.XLS,
XLSTART.XLS, k4.xls, xl5glary.xls, mypersonnel.xls,
Xlscan.xls, laroux.xls, sheet.xls, auto.xls, ssheet.xls
```

### Malicious VBA Module Names (23+) | Tên module VBA độc hại (23+)

```
car, cure, foxz, lalala, laroux, locas, monci, pldt,
program, results, sgv, startup, wendy, vera, binv,
dimon, ecsystem, kinslayer, negs, nocal, privat,
sing, windos, xlstart, k4, xl5glary
```

### Virus Macro Functions (8+) | Hàm macro virus (8+)

```
auto_open, check_files, ck_files, scan_files,
cop, escape, del, back
```

---

## Understanding Scan Results | Hiểu kết quả quét

### Log Messages | Thông điệp log

**[VIRUS]** - Infected file detected | File nhiễm virus được phát hiện
- Example: `[VIRUS] PERSONAL.XLS - Known virus filename`
- Action: File is marked for removal (if SCAN & CLEAN mode)

**[CLEAN]** - File is safe | File an toàn
- Example: `[CLEAN] MyWorkbook.xls`
- Action: No action needed

**[REMOVED]** - Infected file deleted | File nhiễm virus đã xóa
- Example: `[REMOVED] C:\...\XLSTART\VERA.XLS`
- Action: File successfully removed

**[ERROR]** - Could not remove file | Không thể xóa file
- Example: `[ERROR] Could not remove: STARTUP.XLS - File may be in use`
- Action: Close Excel and try again

---

## Troubleshooting | Khắc phục sự cố

### Problem: "Could not remove file - File may be in use"

**English:**
- Close all Excel instances
- Close any programs that might be using Excel files
- Run the cleaner again

**Vietnamese:**
- Đóng tất cả cửa sổ Excel
- Đóng các chương trình có thể đang sử dụng file Excel
- Chạy lại công cụ quét

### Problem: No viruses detected but Excel still behaves strangely

**English:**
1. The virus might be in a workbook file, not XLSTART
2. Use the main Excel Converter to convert and clean your workbooks
3. Check for other malware using Windows Defender or antivirus

**Vietnamese:**
1. Virus có thể nằm trong file workbook, không phải XLSTART
2. Sử dụng Excel Converter chính để chuyển đổi và làm sạch workbook
3. Kiểm tra malware khác bằng Windows Defender hoặc antivirus

### Problem: Tool won't start

**English:**
- Make sure PowerShell is enabled
- Right-click `XLSStartCleaner.ps1` → Run with PowerShell
- Check if Windows Defender is blocking the script

**Vietnamese:**
- Đảm bảo PowerShell được bật
- Nhấp chuột phải `XLSStartCleaner.ps1` → Run with PowerShell
- Kiểm tra xem Windows Defender có chặn script không

---

## Command Line Usage | Sử dụng dòng lệnh

### Silent Mode (No GUI) | Chế độ im lặng (Không GUI)

**Scan only:**
```powershell
powershell.exe -ExecutionPolicy Bypass -File "XLSStartCleaner.ps1" -Silent
```

**Scan and auto-clean:**
```powershell
powershell.exe -ExecutionPolicy Bypass -File "XLSStartCleaner.ps1" -Silent -AutoClean
```

**Exit codes:**
- `0` = No viruses found
- `>0` = Number of infected files found

---

## Safety & Privacy | An toàn & Riêng tư

**What this tool does:**
- ✅ Scans XLSTART folders only
- ✅ Detects known virus patterns
- ✅ Removes infected files (only in CLEAN mode)
- ✅ Runs completely offline (no internet required)

**What this tool does NOT do:**
- ❌ Does not send any data to the internet
- ❌ Does not modify your Excel workbooks
- ❌ Does not access files outside XLSTART folders
- ❌ Does not install anything on your system

---

## References | Tham khảo

- Microsoft Threat Encyclopedia: [X97M/Laroux](https://www.microsoft.com/en-us/wdsi/threats/malware-encyclopedia-description?Name=Virus:X97M/Laroux)
- TrendMicro Threat Database: [X97M_LAROUX](https://www.trendmicro.com/vinfo/us/threat-encyclopedia/malware/X97M_LAROUX.RR)

---

## Support | Hỗ trợ

If you find this tool useful, please consider supporting the developer.

Nếu bạn thấy công cụ này hữu ích, hãy cân nhắc ủng hộ tác giả.

| Method | Account | Name |
|--------|---------|------|
| MB Bank | `0360126996868` | LE VAN AN |
| Momo | `0976896621` | LE VAN AN |

---

## License | Giấy phép

MIT License

Copyright (c) 2026 Le An (Vietnam IT)

---

**Stay safe! | Giữ an toàn!** 🛡️
