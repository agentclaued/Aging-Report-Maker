@echo off
rem Development: stops the background app, runs start_wallmartsummary_app.bat (with --reload),
rem and starts the background app again when you close the dev server.
call "%~dp011 2 OFF - Stop in background.bat"
call "%~dp0..\..\start_wallmartsummary_app.bat"
call "%~dp011 1 ON - Start in background.bat"
