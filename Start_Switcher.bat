@echo off
setlocal
set SCRIPT=%~dp0Switcher.ps1
if not exist "%SCRIPT%" (
    echo Switcher.ps1 not found. Keep both files in the same folder.
    pause & exit /b 1
)

:: Kill ALL previous Switcher instances.
:: Old "WINDOWTITLE" filter misses hidden windows — use CommandLine match instead.
wmic process where "name='powershell.exe' and commandline like '%%Switcher.ps1%%'" delete >nul 2>&1
timeout /t 1 /nobreak >nul

powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "%SCRIPT%"
