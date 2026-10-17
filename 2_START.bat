@echo off
cd /d "%~dp0"
title Restock Watch - keep this window open (minimise it)
set /p PORT=Port number (just press Enter for 8787):
if "%PORT%"=="" set PORT=8787
echo Opening the dashboard in your browser...
python stock_tracker.py web --port %PORT%
pause
