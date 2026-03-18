#!/bin/bash

# 创建离线安装包
# 包含所有必要文件，用户可以在本地运行

echo "正在创建离线安装包..."

PACKAGE_NAME="word-addin-offline.zip"
TEMP_DIR="word-addin-offline-temp"

# 清理
if [ -f "$PACKAGE_NAME" ]; then
    rm "$PACKAGE_NAME"
fi
if [ -d "$TEMP_DIR" ]; then
    rm -rf "$TEMP_DIR"
fi

# 创建临时目录
mkdir -p "$TEMP_DIR"

# 复制必要文件
echo "复制文件..."
cp manifest-local.xml "$TEMP_DIR/"
cp taskpane.html "$TEMP_DIR/"
cp taskpane.css "$TEMP_DIR/"
cp taskpane.js "$TEMP_DIR/"
cp logger.js "$TEMP_DIR/"
cp commands.html "$TEMP_DIR/"
cp dialog.html "$TEMP_DIR/"

# 复制服务器文件
cp start-server.py "$TEMP_DIR/"
cp 启动服务器.command "$TEMP_DIR/"
cp start-server.bat "$TEMP_DIR/"

# 复制 assets 目录（如果存在）
if [ -d "assets" ]; then
    cp -r assets "$TEMP_DIR/"
fi

# 复制背景图片文件（如果存在）
if [ -f "miku-background.jpg" ]; then
    cp miku-background.jpg "$TEMP_DIR/"
fi
# 复制其他可能的图片文件
for img in *.jpg *.png *.gif *.jpeg; do
    if [ -f "$img" ]; then
        cp "$img" "$TEMP_DIR/"
    fi
done

# 复制用户文档
cp USER_OFFLINE_INSTALL.md "$TEMP_DIR/README.md"
cp 故障排除指南.md "$TEMP_DIR/"
cp 快速诊断.txt "$TEMP_DIR/"
cp test-server.html "$TEMP_DIR/"

# 设置执行权限
chmod +x "$TEMP_DIR/启动服务器.command"
chmod +x "$TEMP_DIR/start-server.py"

# 创建 ZIP 包
echo "创建 ZIP 包..."
cd "$TEMP_DIR"
# 包含故障排除指南和快速诊断文件，但排除其他开发文档
zip -r "../$PACKAGE_NAME" . -x "*.log" "node_modules/*" "*.pem" ".gitignore" "*.git/*" "DEPLOYMENT.md" "DISTRIBUTION_GUIDE.md" "OFFLINE_DEPLOYMENT.md" "INSTALL.md" "QUICKSTART.md" "README.md" "create-*.sh" "update-*.js" "generate-*.sh" "server.js" "package.json"
cd ..

# 清理临时目录
rm -rf "$TEMP_DIR"

echo "✓ 离线安装包已创建: $PACKAGE_NAME"
echo ""
echo "这个包包含："
echo "  - 所有插件文件"
echo "  - 本地服务器脚本（Python）"
echo "  - 启动脚本（macOS 和 Windows）"
echo "  - 用户安装指南"
echo ""
echo "用户只需："
echo "  1. 解压文件"
echo "  2. 双击启动服务器"
echo "  3. 在 Word 中加载 manifest-local.xml"

