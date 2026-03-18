@echo off
REM Word Add-in 本地服务器（Windows）— 在项目根目录双击运行
cd /d "%~dp0"

echo ========================================
echo Word Add-in 本地服务器
echo ========================================
echo.

python --version >nul 2>&1
if errorlevel 1 (
    echo 错误: 未找到 Python
    pause
    exit /b 1
)

python scripts\start-server.py
pause
