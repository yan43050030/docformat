@echo off
title DocFormat Pro
cd /d "%~dp0"
python main.py
if errorlevel 1 (
    echo.
    echo Qi Dong Shi Bai: Qing Xian An Zhuang Yi Lai / Startup failed, install deps first:
    echo     pip install -r requirements.txt
    pause
)
