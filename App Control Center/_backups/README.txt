APP CONTROL CENTER - how to run WalMartSummary
Files are grouped by number: each number is one way to run the app. ON starts it, OFF stops/removes it.

00  Terminal      00 1 ON - Start in terminal.bat               Visible terminal, no reload (stops the background app first).
                  00 2 OFF - Stop in terminal.bat              Stops the server.
11  Background    11 1 ON - Start in background.bat             Starts the hidden tray app now.
                  11 2 OFF - Stop in background.bat             Stops the hidden tray app and server.
22  PC boot       22 1 ON - Auto-start at PC boot (no recovery).bat   Run once (Administrator). Server starts ~45s after the PC turns on,
                                                              before login, hidden, once. If it stops or crashes it is NOT restarted.
                                                              Removes the login auto-start (33). Task: WalMartSummary_Boot.
                  22 2 OFF - Remove PC boot auto-start.bat      Removes the boot task. Does not stop a running server.
33  Login         33 1 ON - Auto-start at login (tray, recovery).bat   Run once. Starts hidden at every Windows login (tray icon, restarts on crash).
                  33 2 OFF - Remove login auto-start.bat        Removes the login auto-start.
44  Dev mode      44 1 ON - Dev mode (reload).bat               Edit code with --reload (stops background app, restores it afterwards).
                  44 2 OFF - Stop dev mode.bat                  Stops the server.

Helpers (not for double-click): boot_autostart.ps1, stop_background.ps1, background_tray.pyw (the tray app). Old file copies: _backups\
The original start_wallmartsummary_app.bat / stop_wallmartsummary_app.bat in the project folder are unchanged.
HOST and PORT are read from app_runtime.env (currently port 8040), shared by all modes, so only one
mode runs at a time. After code edits the background app needs "Restart server" from its tray menu.
The boot server has no tray: to restart it use 11 OFF then: schtasks /Run /TN WalMartSummary_Boot
Server log: logs\server.log
Requires pystray and pillow (listed in requirements.txt; the login auto-start installs them if missing).
