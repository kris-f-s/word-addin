#!/bin/bash

# Word Add-in 本地服务器启动脚本（macOS）
# 双击此文件即可启动服务器

# 获取脚本所在目录
SCRIPT_DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )" && pwd )"
cd "$SCRIPT_DIR"

# 设置窗口标题
echo -ne "\033]0;Word Add-in 服务器\007"

# 清屏
clear

# 检查 Python
if ! command -v python3 &> /dev/null; then
    echo "错误: 未找到 Python 3"
    echo "请先安装 Python 3"
    echo ""
    read -p "按回车键退出..."
    exit 1
fi

# 启动 Python 服务器
python3 start-server.py

# 如果服务器退出，保持窗口打开
echo ""
read -p "服务器已停止，按回车键退出..."

