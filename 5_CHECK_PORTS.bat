@echo off
cd /d "%~dp0"
python stock_tracker.py ports --from 8000 --to 9000
echo.
echo ===== All busy ports on this PC (last column = program ID) =====
netstat -ano | findstr LISTENING
pause
