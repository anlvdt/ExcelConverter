@echo off
:: Excel Converter v1.1.0
:: Author: Le An (Vietnam IT)
:: Convert XLS/XLSM to XLSX with virus protection
cd /d "%~dp0"
start "" /b powershell -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -Command "& {Start-Process powershell -ArgumentList '-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File \"%~dp0ConvertXLS_GUI.ps1\"' -WindowStyle Hidden}"
exit
