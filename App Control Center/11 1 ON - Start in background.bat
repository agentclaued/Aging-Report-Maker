@echo off
rem Starts the hidden WalMartSummary tray app (tray icon) now, without waiting for the next login.
pushd "%~dp0.."
for /f "delims=" %%P in ('python -c "import sys, os; print(os.path.join(os.path.dirname(sys.executable), 'pythonw.exe'))"') do set "PYW=%%P"
start "" "%PYW%" "%~dp0background_tray.pyw"
popd
echo Background app started. Look for the icon near the clock.
timeout /t 3 >nul
