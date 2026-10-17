@echo off
cd /d "%~dp0"
echo Installing Stock Tracker...
python -m pip install -r requirements.txt
python -m playwright install chromium
python -c "import stock_tracker; stock_tracker.load_config()"
echo.
echo Done! Now double-click 2_START.bat
pause
