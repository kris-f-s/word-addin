@echo off
REM Word Add-in 本地服务器启动脚本（Windows）
REM 双击此文件即可启动服务器

cd /d "%~dp0"

echo ========================================
echo Word Add-in 本地服务器
echo ========================================
echo.

REM 检查 Python
python --version >nul 2>&1
if errorlevel 1 (
    echo 错误: 未找到 Python
    echo 请先安装 Python 3
    echo.
    pause
    exit /b 1
)

REM 启动 Python 服务器
python start-server.py

pause

