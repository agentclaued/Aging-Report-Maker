@echo off
rem Development: stops the background app, runs start_wallmartsummary_app.bat (with --reload) (tray icon shows status),
rem and starts the background app again when you close the dev server.
call "%~dp011 2 OFF - Stop in background.bat"
pushd "%~dp0.."
for /f "delims=" %%P in ('python -c "import sys, os; print(os.path.join(os.path.dirname(sys.executable), 'pythonw.exe'))"') do set "PYW=%%P"
popd
start "" "%PYW%" "%~dp0background_tray.pyw" --attach
call "%~dp0..\start_wallmartsummary_app.bat"
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0stop_background.ps1"
call "%~dp011 1 ON - Start in background.bat"
