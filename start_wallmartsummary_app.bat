@echo off
setlocal EnableExtensions EnableDelayedExpansion
cd /d "%~dp0"

set "CONFIG_FILE=app_runtime.env"
set "HOST=192.168.2.123"
set "PORT=8040"
set "PYTHON_EXE="

if exist "%CONFIG_FILE%" (
    for /f "usebackq tokens=1,* delims==" %%A in ("%CONFIG_FILE%") do (
        set "KEY=%%A"
        set "VALUE=%%B"
        if /I "!KEY!"=="HOST" set "HOST=!VALUE!"
        if /I "!KEY!"=="PORT" set "PORT=!VALUE!"
        if /I "!KEY!"=="PYTHON_EXE" set "PYTHON_EXE=!VALUE!"
    )
)

if defined PYTHON_EXE (
    if exist "!PYTHON_EXE!" goto run_app
)

for %%P in (
    "%LocalAppData%\Python\pythoncore-3.14-64\python.exe"
    "%LocalAppData%\Programs\Python\Python314\python.exe"
) do (
    if exist %%~P (
        set "PYTHON_EXE=%%~P"
        goto run_app
    )
)

echo Python was not found.
echo Set PYTHON_EXE in app_runtime.env and run this file again.
pause
exit /b 1

:run_app
"!PYTHON_EXE!" -m uvicorn app.main:app --host !HOST! --port !PORT! --reload
endlocal
