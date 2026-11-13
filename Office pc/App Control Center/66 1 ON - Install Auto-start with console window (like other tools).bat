@echo off
rem Same style as the other office tools: boot task that opens a visible console window; closing the window stops the app.
rem Needs Administrator rights (asks automatically). Run 11 2 OFF first so port 8507 is free.
net session >nul 2>&1 || (powershell -NoProfile -Command "Start-Process -Verb RunAs -FilePath '%~f0'" & exit /b)
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0window_autostart.ps1" -Action Install
pause
