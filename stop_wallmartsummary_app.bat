@echo off
setlocal EnableExtensions EnableDelayedExpansion
cd /d "%~dp0"

set "CONFIG_FILE=app_runtime.env"
set "PORT=8040"

if exist "%CONFIG_FILE%" (
    for /f "usebackq tokens=1,* delims==" %%A in ("%CONFIG_FILE%") do (
        if /I "%%A"=="PORT" set "PORT=%%B"
    )
)

REM Stops only the WalMartSummary uvicorn process on this port; other apps are left running.
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0stop_wallmartsummary_app.ps1" -Port !PORT!

endlocal
