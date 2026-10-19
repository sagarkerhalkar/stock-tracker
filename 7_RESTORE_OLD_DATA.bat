@echo off
cd /d "%~dp0"
title Restock Watch - bring back old data
python stock_tracker.py restore
echo.
echo Now double-click 2_START.bat (or press Check now on the dashboard).
pause
