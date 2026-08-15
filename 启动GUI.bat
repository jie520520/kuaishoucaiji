@echo off
chcp 65001 >nul 2>&1
title 快手达人采集工具 GUI
cd /d "%~dp0"

REM 用系统 Python 启动器 py -3（自带 tkinter），避免 workbuddy 精简 python 无 tcl 运行时
where py >nul 2>&1
if errorlevel 1 (
    powershell -NoProfile -Command "Add-Type -AssemblyName System.Windows.Forms; [System.Windows.Forms.MessageBox]::Show('未找到 Python 启动器(py.exe)。\n请安装 Python 3.x 并勾选 tcl/tk and IDLE，或确认 py.exe 在 PATH 中。', '启动失败', 'OK', 'Error')"
    pause
    exit /b 1
)

py -3 gui_scraper.py > gui_run.log 2>&1
set RC=%ERRORLEVEL%

if %RC% neq 0 (
    powershell -NoProfile -Command "$log='gui_run.log'; $e=''; try { $e=Get-Content -Path $log -Raw -ErrorAction Stop } catch { $e='' }; if ([string]::IsNullOrWhiteSpace($e)) { $e='GUI 启动失败，但日志为空（退出码 %RC%）。\n可能 py -3 指向了不完整的 Python，或脚本根本没运行。' }; Add-Type -AssemblyName System.Windows.Forms; [System.Windows.Forms.MessageBox]::Show($e, 'GUI 启动失败（退出码 %RC%）', 'OK', 'Error')"
) else (
    echo [正常结束]
)
pause
