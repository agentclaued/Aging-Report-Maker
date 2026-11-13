@echo off
rem Removes the console-window boot task (66). Close the running window yourself to stop the app.
net session >nul 2>&1 || (powershell -NoProfile -Command "Start-Process -Verb RunAs -FilePath '%~f0'" & exit /b)
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0window_autostart.ps1" -Action Remove
pause
