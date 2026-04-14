========================================
    EXCEL CONVERTER v1.1.0
    XLS/XLSM to XLSX Converter
    with X97M/Laroux Virus Protection
========================================

Author: Le An (Vietnam IT)
Date: January 2026

Original conversion script reference:
https://gist.github.com/gabceb/954418

HOW TO USE:
-----------
Double-click "ExcelConverter.bat" to launch

FEATURES:
---------
[Conversion]
- Convert .xls (Excel 97-2003) to .xlsx
- Convert .xlsm (Macro-enabled) to .xlsx
- Batch convert entire folders
- Include subfolders (recursive)
- Move or delete original files
- Activity log with timestamps

[Security - X97M/Laroux Virus Protection]
- Deep scan VBA modules inside workbooks
- Detect malicious macros by reading code content
- Auto-remove infected VBA modules
- Scan all XLSTART folders (Office 2007-2021, 365)
- Auto-scan XLSTART folders before conversion
- Detect 27+ known virus file variants:
  PERSONAL.XLS, VERA.XLS, PLDT.XLS, STARTUP.XLS,
  BOOK1.XLS, CAR.XLS, CURE.XLS, DIMON.XLS,
  ECSYSTEM.XLS, KINSLAYER.XLS, NEGS.XLS, NOCAL.XLS,
  PRIVAT.XLS, RESULTS.XLS, SGV.XLS, SING.XLS,
  WINDOS.XLS, XLSTART.XLS, BINV.XLS
- Detect 23+ malicious VBA module names:
  laroux, car, cure, foxz, lalala, locas, monci,
  pldt, program, results, sgv, startup, wendy,
  vera, binv, dimon, ecsystem, kinslayer, negs,
  nocal, privat, sing, windos, xlstart
- Detect 8+ virus macro functions:
  auto_open, check_files, ck_files, scan_files,
  cop, escape, del, back
- Final security scan after conversion

VIRUS INFORMATION:
------------------
X97M/Laroux is a family of macro viruses that
spreads using Microsoft Excel spreadsheets.
The virus creates infected files in Excel's
XLSTART folder, which auto-runs when Excel opens.

The virus contains two main macros:
- auto_open: executed when infected file opens
- check_files: infects other spreadsheets

Reference:
- Microsoft Threat Encyclopedia
- TrendMicro Threat Database

REQUIREMENTS:
-------------
- Windows 7/8/10/11
- Microsoft Excel installed
- PowerShell (included in Windows)

NOTES:
------
- Original files can be moved to 'old' folder
- Use "Delete original files" with caution!
- Excel runs in background during conversion
- Close Excel before running for best results

========================================
(c) 2026 Le An - Vietnam IT
========================================
