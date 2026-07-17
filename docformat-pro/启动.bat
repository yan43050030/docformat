@echo off
chcp 65001 >nul
title DocFormat Pro - 公文格式自动排版工具

cd /d "%~dp0"

echo.
echo   ╔══════════════════════════════════════╗
echo   ║     DocFormat Pro  v1.0.0           ║
echo   ║  公文格式自动排版工具               ║
echo   ║  GB/T 9704-2012 标准                ║
echo   ╚══════════════════════════════════════╝
echo.
echo   正在启动服务...

start "" http://127.0.0.1:5173

call npm run dev

echo.
echo   服务已停止。
pause
