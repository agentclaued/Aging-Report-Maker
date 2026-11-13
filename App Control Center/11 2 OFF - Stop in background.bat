@echo off
rem Stops this project's background tray app and its hidden server only (other apps and dev windows are not touched).
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0stop_background.ps1"
echo Background app stopped.
timeout /t 2 >nul
