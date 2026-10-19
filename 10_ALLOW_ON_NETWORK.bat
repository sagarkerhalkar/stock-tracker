@echo off
cd /d "%~dp0"
title Restock Watch - allow phones and PCs on your network
net session >nul 2>&1
if %errorlevel% neq 0 (
  echo  Asking Windows for permission...
  powershell -NoProfile -Command "Start-Process -FilePath '%~f0' -Verb RunAs"
  exit /b
)
netsh advfirewall firewall delete rule name="Restock Watch 8787" >nul 2>&1
netsh advfirewall firewall add rule name="Restock Watch 8787" dir=in action=allow protocol=TCP localport=8787 profile=private,domain
echo.
echo  Done. Phones and PCs on your home or office network can now open Restock Watch at:
for /f "tokens=2 delims=:" %%a in ('ipconfig ^| findstr /c:"IPv4"') do echo     http:%%a:8787
echo.
echo  (Public Wi-Fi networks stay blocked. If your home network is set to "Public" in Windows,
echo   change it to "Private": Settings - Network and Internet - your network - Private.)
echo.
pause
