# v1.2.0 — 2026-10-26

- Fix queue deduplication losing Name/Directory/BaseName and breaking output paths.
- Add native Unicode long-path IO and short local staging for OneDrive/SharePoint synced folders; XLSB support.
- Report existing XLSX explicitly; archives must be extracted first.
- Isolate each file in a worker with timeout/cancel and owned Excel process cleanup.
- Validate/reopen output, verify destination hash, publish without overwriting, preserve originals by default.
- Remove filename-based deletion and automatic XLSTART cleanup from conversion.
- Add per-file CSV reports, headless batch entry point and regression/Windows Excel smoke tests.
- External-link workbooks retain originals with a review warning.

Verification: 12 core tests, 6 subprocess tests, and script parsing pass on macOS PowerShell 7 and Windows CI with PowerShell 5.1 and 7. Actual desktop Excel COM, WinForms, and live OneDrive/SharePoint sync integration remain to be run; see the audit report.

---

## Excel Converter v1.1.0

**Convert XLS/XLSM to XLSX with X97M/Laroux Virus Protection**

### Features | Tính năng

**Conversion | Chuyển đổi:**
- Convert XLS (Excel 97-2003) to XLSX
- Convert XLSM (Macro-enabled) to XLSX
- Batch convert folders with subfolders
- Move or delete original files

**Security | Bảo mật:**
- Deep scan VBA modules inside workbooks
- Detect malicious macros by reading code content
- Auto-remove infected VBA modules
- Scan all XLSTART folders (Office 2007-2021, 365)
- Detect 27+ virus file variants
- Detect 23+ malicious module names
- Detect 8+ virus macro functions
- **NEW: XLSStart Cleaner** - Standalone tool to scan/clean XLSTART folders without conversion
- **NEW: Scan XLSTART button** - Quick access to virus scanner from main GUI

### Installation | Cài đặt
1. Download ExcelConverter-v1.1.0.zip
2. Extract the ZIP file
3. Double-click ExcelConverter.cmd to run

### Requirements | Yêu cầu
- Windows 7/8/10/11
- Microsoft Excel installed
- PowerShell (included in Windows)

### Donate | Ủng hộ
| Method | Account | Name |
|--------|---------|------|
| MB Bank | 0360126996868 | LE VAN AN |
| Momo | 0976896621 | LE VAN AN |

---
Author: Le An (Vietnam IT)
