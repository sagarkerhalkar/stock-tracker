@echo off
cd /d "%~dp0"
title Restock Watch - http://localhost:8787 - keep this window open (minimise it)
if /i "%~1"=="silent" set NO_BROWSER=1

:run
echo.
echo  Starting Restock Watch on http://localhost:8787 ...
python stock_tracker.py web
set CODE=%errorlevel%
if "%CODE%"=="0" goto stopped
if "%CODE%"=="3" goto portbusy
if "%CODE%"=="9009" goto nopython
echo.
echo  Restock Watch stopped unexpectedly - code %CODE%. Restarting in 10 seconds.
echo  Close this window if you want to stop it.
timeout /t 10 /nobreak >nul
goto run

:portbusy
echo  Waiting 60 seconds for port 8787 to become free, then trying again.
echo  Close this window if you want to stop it.
timeout /t 60 /nobreak >nul
goto run

:nopython
echo.
echo  Python was not found. Install Python 3.10+ from python.org - tick "Add Python to PATH" - then run 1_INSTALL.bat.
pause
exit /b 1

:stopped
echo.
echo  Restock Watch has stopped.
if /i not "%~1"=="silent" pause
exit /b 0
