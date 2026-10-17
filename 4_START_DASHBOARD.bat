@echo off
cd /d "%~dp0"
title Stock Tracker Dashboard - keep this window open
set /p PORT=Port number (press Enter for 8787):
if "%PORT%"=="" set PORT=8787
python stock_tracker.py web --port %PORT%
pause
