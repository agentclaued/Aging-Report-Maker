@echo off
rem Removes the PC-restart auto-start (55). Does not stop a running server (use 11 2 OFF).
net session >nul 2>&1 || (powershell -NoProfile -Command "Start-Process -Verb RunAs -FilePath '%~f0'" & exit /b)
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0boot_recovery_autostart.ps1" -Action Remove
pause
