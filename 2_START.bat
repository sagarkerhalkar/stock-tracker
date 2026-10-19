@echo off
cd /d "%~dp0"
title Restock Watch - http://localhost:8787 - keep this window open (minimise it)
if /i "%~1"=="silent" set NO_BROWSER=1
echo Starting Restock Watch on http://localhost:8787 ...
python stock_tracker.py web
echo.
echo Restock Watch has stopped. Read the message above, then close this window.
pause
