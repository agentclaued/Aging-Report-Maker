@echo off
rem Runs WalMartSummary every time the PC starts or restarts (before login, hidden) and restarts it if it stops.
rem Needs Administrator rights (asks automatically). Replaces the 22 and 33 auto-starts.
net session >nul 2>&1 || (powershell -NoProfile -Command "Start-Process -Verb RunAs -FilePath '%~f0'" & exit /b)
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0boot_recovery_autostart.ps1" -Action Install
pause
