#!/bin/bash
# 创建用户安装包（仅 manifest.xml）
set -e
ROOT="$(cd "$(dirname "$0")/.." && pwd)"
cd "$ROOT"

PACKAGE_NAME="word-addin-install.zip"
echo "正在创建用户安装包..."

if [ -f "$PACKAGE_NAME" ]; then
    rm "$PACKAGE_NAME"
fi

zip "$PACKAGE_NAME" manifest.xml

echo "✓ 安装包已创建: $PACKAGE_NAME"
