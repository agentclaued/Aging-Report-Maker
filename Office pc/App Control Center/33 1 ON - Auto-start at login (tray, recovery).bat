@echo off
rem Starts WalMartSummary in the background (tray icon) every time you log in to Windows.
rem Creates a Startup-folder shortcut (no admin needed) and starts the tray app now.
pushd "%~dp0.."

where python >nul 2>nul
if errorlevel 1 (
    echo Python was not found on PATH. Install Python 3.11+ from https://python.org and try again.
    pause
    exit /b 1
)

python -c "import fastapi, uvicorn, pystray, PIL, pandas, openpyxl, pyodbc" 2>nul
if errorlevel 1 (
    echo Installing required packages...
    python -m pip install -r requirements.txt
    if errorlevel 1 (
        echo Could not install dependencies. Check your internet connection and try again.
        pause
        exit /b 1
    )
)

for /f "delims=" %%P in ('python -c "import sys, os; print(os.path.join(os.path.dirname(sys.executable), 'pythonw.exe'))"') do set "PYW=%%P"
if not exist "%PYW%" (
    echo pythonw.exe not found at %PYW%
    pause
    exit /b 1
)

set "LNK=%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup\WalMartSummary.lnk"
set "SCRIPT=%~dp0background_tray.pyw"
set "ROOT=%CD%"

powershell -NoProfile -ExecutionPolicy Bypass -Command "$s=(New-Object -ComObject WScript.Shell).CreateShortcut($env:LNK); $s.TargetPath=$env:PYW; $s.Arguments='\"'+$env:SCRIPT+'\"'; $s.WorkingDirectory=$env:ROOT; $s.WindowStyle=7; $s.Description='WalMartSummary (background)'; $s.Save()"
if errorlevel 1 (
    echo Could not create the startup shortcut.
    pause
    exit /b 1
)

echo.
echo Done. WalMartSummary now starts automatically when you log in to Windows.
echo Look for the WalMartSummary icon near the clock (it may be under the ^ arrow).
echo To remove auto-start, run 33 2 OFF - Remove login auto-start.bat.
echo.
start "" "%PYW%" "%SCRIPT%"
popd
timeout /t 3 >nul
