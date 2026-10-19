@echo off
cd /d "%~dp0"
title Restock Watch - start with Windows
echo.
echo  Restock Watch starts with Windows automatically (this is ON by default).
echo.
echo    1 = keep it ON
echo    2 = turn it OFF
echo.
set /p CH=Type 1 or 2 and press Enter: 
if "%CH%"=="2" (python stock_tracker.py autostart off) else (python stock_tracker.py autostart on)
echo.
pause
