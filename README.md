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
- Detect 19+ known virus file variants | Phát hiện 19+ biến thể virus
- Detect 21+ malicious VBA module names | Phát hiện 21+ tên module độc hại
- Detect 8+ virus macro functions | Phát hiện 8+ hàm macro virus
- Final security scan after conversion | Quét bảo mật sau khi chuyển đổi

---

## Installation | Cài đặt

1. Download the latest release | Tải phiên bản mới nhất
2. Extract the ZIP file | Giải nén file ZIP
3. Double-click `ExcelConverter.cmd` to run | Double-click `ExcelConverter.cmd` để chạy

---

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
