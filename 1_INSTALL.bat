@echo off
cd /d "%~dp0"
echo Installing Stock Tracker...
python -m pip install -r requirements.txt
python -m playwright install chromium
python -c "import stock_tracker; stock_tracker.load_config()"
python stock_tracker.py autostart on
echo.
echo Windows will now ask permission so phones on your network can open Restock Watch - click Yes.
call 10_ALLOW_ON_NETWORK.bat
echo.
echo Done! Restock Watch will now also start by itself whenever Windows starts.
echo Now double-click 2_START.bat - your dashboard is http://localhost:8787
pause
