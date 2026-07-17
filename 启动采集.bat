@echo off
chcp 65001 >nul
title 颜阿娇 - 快手达人采集工具

cd /d "%~dp0"

set PYTHON=C:\Users\Administrator\.workbuddy\binaries\python\envs\default\Scripts\python.exe

echo ========================================
echo   颜阿娇 - 快手达人采集工具
echo ========================================
echo.
echo   请选择采集模式：
echo.
echo   [1] 推荐标签 (健康/运动/美妆/穿搭... 共8个)
echo   [2] 全部37个标签 (数据量大，耗时较长)
echo   [3] 指定标签 (手动输入)
echo   [4] 交互模式 (菜单选择)
echo   [5] 预览 (不实际采集)
echo   [0] 退出
echo.

set /p choice="请输入数字: "

if "%choice%"=="1" %PYTHON% "%~dp0yanajiao_scraper.py" --rec
if "%choice%"=="2" %PYTHON% "%~dp0yanajiao_scraper.py" --all
if "%choice%"=="3" call :run_tag
if "%choice%"=="4" %PYTHON% "%~dp0yanajiao_scraper.py"
if "%choice%"=="5" %PYTHON% "%~dp0yanajiao_scraper.py" --rec --dry-run

echo.
echo 按任意键关闭...
pause >nul
exit /b

:run_tag
set /p tags="请输入标签名(逗号分隔): "
%PYTHON% "%~dp0yanajiao_scraper.py" --tag %tags%
goto :eof
