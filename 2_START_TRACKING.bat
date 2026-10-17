@echo off
cd /d "%~dp0"
title Stock Tracker - keep this window open
python stock_tracker.py run
pause
