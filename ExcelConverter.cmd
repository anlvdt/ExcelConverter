@echo off
setlocal
:: Use the single launcher for consistent STA and path quoting.
powershell.exe -NoProfile -STA -ExecutionPolicy Bypass -WindowStyle Hidden -File "%~dp0ConvertXLS_GUI.ps1"
endlocal
