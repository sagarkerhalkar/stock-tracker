@echo off
cd /d "%~dp0"
title Restock Watch - http://localhost:8787 - keep this window open (minimise it)
where python >nul 2>&1
if errorlevel 1 goto nopython
python stock_tracker.py launcher %*
if /i not "%~1"=="silent" pause
exit /b 0
:nopython
echo.
echo  Python was not found. Install Python 3.10+ from python.org - tick "Add Python to PATH" - then run 1_INSTALL.bat.
pause
exit /b 1
