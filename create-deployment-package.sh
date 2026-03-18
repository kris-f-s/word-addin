#!/bin/bash

# 创建完整部署包
# 包含所有必要文件（用于本地服务器部署）

echo "正在创建部署包..."

PACKAGE_NAME="word-addin-package.zip"
TEMP_DIR="word-addin-temp"

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
cp manifest.xml "$TEMP_DIR/"
cp taskpane.html "$TEMP_DIR/"
cp taskpane.css "$TEMP_DIR/"
cp taskpane.js "$TEMP_DIR/"
cp commands.html "$TEMP_DIR/"
cp dialog.html "$TEMP_DIR/"

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

# 复制服务器文件（可选）
if [ -f "server.js" ]; then
    cp server.js "$TEMP_DIR/"
fi
if [ -f "package.json" ]; then
    cp package.json "$TEMP_DIR/"
fi
if [ -f "generate-cert.sh" ]; then
    cp generate-cert.sh "$TEMP_DIR/"
    chmod +x "$TEMP_DIR/generate-cert.sh"
fi

# 创建 README
cat > "$TEMP_DIR/README.txt" << EOF
Word 文本颜色格式化插件 - 部署包

安装说明：
1. 解压此文件
2. 按照 USER_DEPLOY.md 中的说明进行部署
3. 或联系 IT 支持获取帮助

文件说明：
- manifest.xml: 插件清单文件
- taskpane.*: 插件界面文件
- server.js: Node.js 服务器（可选）
- assets/: 图标文件

更多信息请参考文档。
EOF

# 创建 ZIP 包
echo "创建 ZIP 包..."
cd "$TEMP_DIR"
zip -r "../$PACKAGE_NAME" . -x "*.md" "*.sh" "*.log" "node_modules/*" "*.pem" ".gitignore" "*.git/*"
cd ..

# 清理临时目录
rm -rf "$TEMP_DIR"

echo "✓ 部署包已创建: $PACKAGE_NAME"
echo ""
echo "这个包包含所有必要文件，用户可以："
echo "1. 解压文件"
echo "2. 部署到 Web 服务器"
echo "3. 在 Word 中加载 manifest.xml"

