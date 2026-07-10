@echo off
REM ============================================================
REM  run_refresh.bat
REM  This is what Windows Task Scheduler runs each day.
REM  It uses the project's private Python (the venv) so the
REM  installed libraries never clash with anything else.
REM ============================================================

cd /d "C:\Projects\Weather_Auto"
"C:\Projects\Weather_Auto\venv\Scripts\python.exe" "C:\Projects\Weather_Auto\refresh_weather.py"
