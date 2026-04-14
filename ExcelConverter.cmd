@echo off
:: Excel Converter v1.1.0
:: Author: Le An (Vietnam IT)
:: Convert XLS/XLSM to XLSX with virus protection

if "%~1"=="hidden" goto :run
start "" /min cmd /c "%~f0" hidden
exit

:run
cd /d "%~dp0"
powershell -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "%~dp0ConvertXLS_GUI.ps1"
