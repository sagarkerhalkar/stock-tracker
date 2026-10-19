@echo off
cd /d "%~dp0"
title Restock Watch - Cloudflare quick tunnel
echo.
echo  Makes a temporary internet link to your Restock Watch (http://localhost:8787).
echo  Anyone with the link sees the sign-in page - your Restock Watch password is required.
echo.
where cloudflared >nul 2>&1
if %errorlevel%==0 goto run
echo  Installing cloudflared from Cloudflare...
winget install --id Cloudflare.cloudflared -e --accept-source-agreements --accept-package-agreements
where cloudflared >nul 2>&1
if %errorlevel%==0 goto run
if exist "%ProgramFiles(x86)%\cloudflared\cloudflared.exe" set "PATH=%PATH%;%ProgramFiles(x86)%\cloudflared"
if exist "%ProgramFiles%\cloudflared\cloudflared.exe" set "PATH=%PATH%;%ProgramFiles%\cloudflared"
where cloudflared >nul 2>&1
if %errorlevel%==0 goto run
echo.
echo  Could not install cloudflared automatically. Close this window, open a NEW one and run this file again,
echo  or download cloudflared for Windows from Cloudflare's website.
pause
exit /b 1
:run
echo.
echo  Look below for a line with  https://....trycloudflare.com  - that is your link.
echo  Keep this window open while you use the link. Close it to switch the link off.
echo.
cloudflared tunnel --url http://localhost:8787
pause
