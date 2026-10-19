@echo off
cd /d "%~dp0"
title Restock Watch - start with Windows
set "LNK=%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup\Restock Watch.lnk"
if exist "%LNK%" goto remove
powershell -NoProfile -ExecutionPolicy Bypass -Command "$s=(New-Object -ComObject WScript.Shell).CreateShortcut($env:LNK); $s.TargetPath='%~dp02_START.bat'; $s.Arguments='silent'; $s.WorkingDirectory='%~dp0'; $s.WindowStyle=7; $s.Save()"
if errorlevel 1 goto failed
echo.
echo  ON: Restock Watch now starts by itself (minimised) every time you sign in to Windows.
echo  Your dashboard is always at  http://localhost:8787  - bookmark it.
echo  Run this file again to turn it off.
echo.
pause
exit /b 0
:remove
del "%LNK%"
echo.
echo  OFF: Restock Watch will no longer start with Windows.
echo.
pause
exit /b 0
:failed
echo  Could not create the shortcut. You can add it by hand: Win+R, type shell:startup, and put a shortcut to 2_START.bat there.
pause
exit /b 1
