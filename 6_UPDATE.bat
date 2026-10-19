@echo off
setlocal
cd /d "%~dp0"
title Restock Watch - update
echo.
echo  Restock Watch updater - your products, alerts and history are kept.
echo.
echo  Step 1 of 3: Backing up your products and history...
for /f %%i in ('powershell -NoProfile -Command "Get-Date -Format yyyy-MM-dd_HHmm"') do set STAMP=%%i
set "BK=backup\%STAMP%"
mkdir "%BK%" >nul 2>&1
if exist config.yaml copy /y config.yaml "%BK%\" >nul
if exist data xcopy /e /i /y /q data "%BK%\data" >nul
echo     Backup saved in %BK%
echo.
echo  Step 2 of 3: Getting the latest version from GitHub...
if exist .git goto gitpull
powershell -NoProfile -ExecutionPolicy Bypass -Command "$ErrorActionPreference='Stop'; $z=Join-Path $env:TEMP 'restock_update.zip'; $d=Join-Path $env:TEMP 'restock_update'; Invoke-WebRequest -UseBasicParsing 'https://github.com/sagarkerhalkar/stock-tracker/archive/refs/heads/main.zip' -OutFile $z; if (Test-Path $d) { Remove-Item $d -Recurse -Force }; Expand-Archive $z $d; $src=(Get-ChildItem $d -Directory | Select-Object -First 1).FullName; Get-ChildItem $src -File | Where-Object { $_.Name -ne 'config.yaml' } | Copy-Item -Destination . -Force"
if errorlevel 1 goto failed
goto install

:gitpull
git pull
if errorlevel 1 goto failed

:install
echo.
echo  Step 3 of 3: Installing anything new...
python -m pip install -q -r requirements.txt
echo.
echo  Done! Double-click 2_START.bat - it closes the old version by itself.
echo.
pause
exit /b 0

:failed
echo.
echo  Could not download the update. Nothing was changed - your data is safe.
echo  If your GitHub repository is private, download the update zip and unzip it into this folder instead.
echo.
pause
exit /b 1
