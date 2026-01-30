# Excel Converter v1.0.0

**Convert XLS/XLSM to XLSX with X97M/Laroux Virus Protection**

Author: Le An (Vietnam IT)

---

## Features | Tính năng

### Conversion | Chuyển đổi
- Convert XLS (Excel 97-2003) to XLSX | Chuyển đổi XLS sang XLSX
- Convert XLSM (Macro-enabled) to XLSX | Chuyển đổi XLSM sang XLSX
- Batch convert entire folders with subfolders | Chuyển đổi hàng loạt cả thư mục con
- Move or delete original files after conversion | Di chuyển hoặc xóa file gốc sau khi chuyển đổi

### Security - X97M/Laroux Virus Protection | Bảo mật - Chống virus X97M/Laroux
- Deep scan VBA modules inside workbooks | Quét sâu VBA modules trong workbook
- Detect malicious macros by reading code content | Phát hiện macro độc hại bằng cách đọc code
- Auto-remove infected VBA modules | Tự động xóa VBA modules bị nhiễm
- Scan all XLSTART folders (Office 2007-2021, 365) | Quét tất cả thư mục XLSTART
- Detect 27+ known virus file variants | Phat hien 27+ bien the virus
- Detect 23+ malicious VBA module names | Phat hien 23+ ten module doc hai
- Detect 8+ virus macro functions | Phat hien 8+ ham macro virus
- Final security scan after conversion | Quét bảo mật sau khi chuyển đổi

---

## Installation | Cài đặt

1. Download the latest release | Tải phiên bản mới nhất
2. Extract the ZIP file | Giải nén file ZIP
3. Double-click `ExcelConverter.cmd` to run | Double-click `ExcelConverter.cmd` để chạy

---

## Usage Guide | Hướng dẫn sử dụng

### Step 1: Select Source Folder | Bước 1: Chọn thư mục nguồn
- Click **Browse...** button to select folder containing XLS/XLSM files
- Nhấn nút **Browse...** để chọn thư mục chứa file XLS/XLSM

### Step 2: Configure Options | Bước 2: Cấu hình tùy chọn

| Option | Description | Mô tả |
|--------|-------------|-------|
| **Move original files to 'old' folder** | Move converted files to 'old' subfolder | Di chuyển file gốc vào thư mục con 'old' |
| **Include subfolders** | Scan and convert files in all subfolders | Quét và chuyển đổi file trong tất cả thư mục con |
| **Show Excel window** | Display Excel during conversion (for debugging) | Hiển thị Excel khi chuyển đổi (để debug) |
| **Delete original files** | Permanently delete original files (use with caution!) | Xóa vĩnh viễn file gốc (cẩn thận!) |

### Step 3: Convert | Bước 3: Chuyển đổi
- Click **CONVERT** button to start conversion
- Nhấn nút **CONVERT** để bắt đầu chuyển đổi

### Step 4: Review Results | Bước 4: Xem kết quả
- Check the Activity Log for conversion status and any virus detections
- Kiểm tra Activity Log để xem trạng thái chuyển đổi và phát hiện virus

### About Button (?) | Nút Giới thiệu (?)
- Click **?** button in the header to view features and donate information
- Nhấn nút **?** ở header để xem tính năng và thông tin ủng hộ

### Scan XLSTART Button | Nut Quet XLSTART
- Click **Scan XLSTART** button to launch standalone virus scanner
- Nhan nut **Scan XLSTART** de mo cong cu quet virus doc lap
- This tool scans and removes viruses from Excel XLSTART folders without converting files
- Cong cu nay quet va xoa virus tu thu muc XLSTART ma khong can chuyen doi file

---

## XLSStart Cleaner (Standalone Tool) | Công cụ Quét XLSTART (Độc lập)

### What is XLSStart Cleaner? | XLSStart Cleaner là gì?
A dedicated tool to scan and remove X97M/Laroux virus from Excel XLSTART folders without converting files.

Công cụ chuyên dụng để quét và xóa virus X97M/Laroux từ thư mục XLSTART mà không cần chuyển đổi file.

### How to use | Cách sử dụng:

**Method 1: From Main GUI | Phuong phap 1: Tu GUI chinh**
1. Open Excel Converter | Mo Excel Converter
2. Click **Scan XLSTART** button in the header | Nhan nut **Scan XLSTART** o header
3. XLSStart Cleaner will open in a new window | XLSStart Cleaner se mo trong cua so moi

**Method 2: Run Standalone | Phương pháp 2: Chạy độc lập**
1. Double-click `XLSStartCleaner.bat` | Double-click `XLSStartCleaner.bat`
2. Or run `XLSStartCleaner.ps1` directly | Hoặc chạy `XLSStartCleaner.ps1` trực tiếp

### Features | Tính năng:
- **SCAN ONLY**: Detect viruses without removing them | Phát hiện virus mà không xóa
- **SCAN & CLEAN**: Detect and permanently remove infected files | Phát hiện và xóa vĩnh viễn file nhiễm virus
- Scans 13+ XLSTART locations (Office 2007-365) | Quét 13+ vị trí XLSTART
- Deep VBA code analysis | Phân tích sâu VBA code
- Detailed scan report | Báo cáo quét chi tiết


## Requirements | Yêu cầu

- Windows 7/8/10/11
- Microsoft Excel installed | Microsoft Excel đã cài đặt
- PowerShell (included in Windows) | PowerShell (có sẵn trong Windows)

---

## Screenshot | Ảnh chụp màn hình

![Excel Converter](screenshot.png)

---

## Virus Detection | Phát hiện Virus

### Known virus file names | Tên file virus đã biết:
```
BINV.XLS, BOOK1.XLS, CAR.XLS, CURE.XLS, DIMON.XLS,
ECSYSTEM.XLS, KINSLAYER.XLS, NEGS.XLS, NOCAL.XLS,
PERSONAL.XLS, PLDT.XLS, PRIVAT.XLS, RESULTS.XLS,
SGV.XLS, SING.XLS, STARTUP.XLS, VERA.XLS, WINDOS.XLS, XLSTART.XLS
```

### Known malicious module names | Tên module độc hại:
```
car, cure, foxz, lalala, laroux, locas, monci, pldt,
program, results, sgv, startup, wendy, vera, binv,
dimon, ecsystem, kinslayer, negs, nocal, privat, sing, windos, xlstart
```

### Known virus macro functions | Hàm macro virus:
```
auto_open, check_files, ck_files, scan_files, cop, escape, del, back
```

---

## References | Tham khảo

- Original conversion script: [gist.github.com/gabceb/954418](https://gist.github.com/gabceb/954418)
- Microsoft Threat Encyclopedia: [X97M/Laroux](https://www.microsoft.com/en-us/wdsi/threats/malware-encyclopedia-description?Name=Virus:X97M/Laroux)
- TrendMicro Threat Database: [X97M_LAROUX](https://www.trendmicro.com/vinfo/us/threat-encyclopedia/malware/X97M_LAROUX.RR)

---

## Donate | Ủng hộ

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

## Changelog | Lịch sử thay đổi

### v1.0.0 (January 2026)
- Initial release | Phiên bản đầu tiên
- XLS/XLSM to XLSX conversion | Chuyển đổi XLS/XLSM sang XLSX
- X97M/Laroux virus protection | Bảo vệ virus X97M/Laroux
- Deep VBA module scanning | Quét sâu VBA module
- Batch folder conversion | Chuyển đổi hàng loạt thư mục
