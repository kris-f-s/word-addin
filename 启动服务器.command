#!/bin/bash
# Word Add-in 本地服务器（macOS）— 在项目根目录双击运行

ROOT="$( cd "$( dirname "${BASH_SOURCE[0]}" )" && pwd )"
cd "$ROOT"

echo -ne "\033]0;Word Add-in 服务器\007"
clear

if ! command -v python3 &> /dev/null; then
    echo "错误: 未找到 Python 3"
    read -p "按回车键退出..."
    exit 1
fi

python3 scripts/start-server.py
echo ""
read -p "服务器已停止，按回车键退出..."
