@echo off
setlocal

powershell.exe -ExecutionPolicy Bypass -NoProfile -NonInteractive -File "%~dp0csv2clgp.ps1"

echo.
echo.
echo.

pause