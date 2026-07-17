@echo off
chcp 65001 >nul 2>&1
title 颜阿娇 - 快手达人采集工具 GUI
cd /d "C:\Users\Administrator\Desktop\颜阿娇\快手达人采集工具"
C:\Users\Administrator\.workbuddy\binaries\python\envs\default\Scripts\python.exe gui_scraper.py
if %ERRORLEVEL% neq 0 (
    echo.
    echo [错误] 程序异常退出，错误代码: %ERRORLEVEL%
    pause
)
