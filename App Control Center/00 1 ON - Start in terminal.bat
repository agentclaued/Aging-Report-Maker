@echo off
rem Runs the server visibly in a terminal WITHOUT --reload (stops the background app first).
call "%~dp011 2 OFF - Stop in background.bat"
pushd "%~dp0.."
for /f "delims=" %%P in ('python -c "import sys, os; print(os.path.join(os.path.dirname(sys.executable), 'pythonw.exe'))"') do set "PYW=%%P"
start "" "%PYW%" "%~dp0background_tray.pyw" --attach
set "HOST=127.0.0.1"
set "PORT=8040"
if exist app_runtime.env (
    for /f "usebackq tokens=1,* delims==" %%A in ("app_runtime.env") do (
        if /I "%%A"=="HOST" set "HOST=%%B"
        if /I "%%A"=="PORT" set "PORT=%%B"
    )
)
python -m uvicorn app.main:app --host %HOST% --port %PORT%
popd
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0stop_background.ps1"
echo Stopped. Run 11 1 ON - Start in background.bat to bring the background app back.
pause
