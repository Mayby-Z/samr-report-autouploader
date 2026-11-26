@echo off
REM 启动「检验检测报告编号上传系统登录自动助手」GUI 版本
REM 用法：双击此文件，或�?PowerShell 中执行：.\run_gui.bat

cd /d "%~dp0"

if not exist auto_upload_cdp_gui.py (
    echo [错误] 找不�?auto_upload_cdp_gui.py，请把此脚本放在项目根目录下�?
    pause
    exit /b 1
)

echo [info] 启动自动登录助手...
python auto_upload_cdp_gui.py

if errorlevel 1 (
    echo.
    echo [提示] 窗口可能已经关闭。如果仍想继续查看日志，请检查上方输出窗口�?
)

pause
