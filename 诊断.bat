@echo off
chcp 65001 >nul 2>&1
title 环境诊断 - 快手达人采集工具
cd /d "%~dp0"

echo ============================================
echo  双击此文件，把下面所有输出抄给开发者
echo ============================================
echo.

echo [1] Python 版本:
py -3 --version 2>&1
echo.

echo [2] tkinter (GUI 库):
py -3 -c "import tkinter; print('  tkinter OK, 版本', tkinter.TkVersion)" 2>&1
echo.

echo [3] yanajiao_scraper (主采集模块):
py -3 -c "import yanajiao_scraper; print('  yanajiao_scraper OK')" 2>&1
echo.

echo [4] extract_contacts (联系方式模块):
py -3 -c "import extract_contacts; print('  extract_contacts OK')" 2>&1
echo.

echo [5] playwright + openpyxl (运行依赖):
py -3 -c "import playwright, openpyxl; print('  playwright + openpyxl OK')" 2>&1
echo.

echo [6] 直接尝试启动 GUI (10秒):
echo   （若下方出现红色窗口/报错，把内容抄给我）
start "" /wait py -3 gui_scraper.py
echo   GUI 进程已结束，退出代码见上。
echo.

echo ============================================
echo  以上全部内容请截图或抄给开发者
echo ============================================
pause
