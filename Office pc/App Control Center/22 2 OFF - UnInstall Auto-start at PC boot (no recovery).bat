@echo off
rem Removes the PC-boot auto-start task. Does not stop a server that is already running.
net session >nul 2>&1 || (powershell -NoProfile -Command "Start-Process -Verb RunAs -FilePath '%~f0'" & exit /b)
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0boot_autostart.ps1" -Action Remove
pause
