@echo off
rem Stops this project's server (also stops one started from a terminal or dev mode). Other apps are not touched.
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0stop_background.ps1"
echo Stopped.
timeout /t 2 >nul
