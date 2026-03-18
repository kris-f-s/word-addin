#!/bin/bash

# 创建用户安装包
# 只包含 manifest.xml（用于从 Web 服务器安装）

echo "正在创建用户安装包..."

PACKAGE_NAME="word-addin-install.zip"

# 清理旧包
if [ -f "$PACKAGE_NAME" ]; then
    rm "$PACKAGE_NAME"
fi

# 创建 ZIP 包，只包含 manifest.xml
zip "$PACKAGE_NAME" manifest.xml

echo "✓ 安装包已创建: $PACKAGE_NAME"
echo ""
echo "这个包包含 manifest.xml，用户可以："
echo "1. 解压文件"
echo "2. 在 Word 中加载 manifest.xml"
echo "3. 插件将从 Web 服务器加载（需要在 manifest.xml 中配置正确的 URL）"

