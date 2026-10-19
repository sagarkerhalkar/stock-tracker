@echo off
cd /d "%~dp0"
echo Who is using port 8787 (Restock Watch's port)?
echo.
netstat -ano | findstr ":8787 " | findstr LISTENING
echo.
echo The last number is the program ID (PID). To see its name:   tasklist /FI "PID eq <number>"
echo If it is not Restock Watch, close that program (or restart the PC) and start 2_START.bat again.
echo.
pause
