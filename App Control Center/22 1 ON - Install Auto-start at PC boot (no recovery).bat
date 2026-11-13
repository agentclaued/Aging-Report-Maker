@echo off
rem Starts WalMartSummary when the PC turns on (before login). Starts once, NO auto-recovery.
rem Needs Administrator rights (asks automatically).
net session >nul 2>&1 || (powershell -NoProfile -Command "Start-Process -Verb RunAs -FilePath '%~f0'" & exit /b)
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0boot_autostart.ps1" -Action Install
pause
