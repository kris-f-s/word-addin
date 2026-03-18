#!/bin/bash
# 离线安装包：扁平目录 + 根目录 start-server.py（与旧版用户流程一致）
set -e
ROOT="$(cd "$(dirname "$0")/.." && pwd)"
cd "$ROOT"

PACKAGE_NAME="word-addin-offline.zip"
TEMP_DIR="word-addin-offline-temp"

if [ -f "$PACKAGE_NAME" ]; then rm "$PACKAGE_NAME"; fi
if [ -d "$TEMP_DIR" ]; then rm -rf "$TEMP_DIR"; fi

mkdir -p "$TEMP_DIR"
echo "复制文件..."
cp manifest-local.xml "$TEMP_DIR/"
cp "$ROOT/public/"*.html "$TEMP_DIR/"
cp "$ROOT/public/taskpane.css" "$ROOT/public/taskpane.js" "$ROOT/public/logger.js" "$TEMP_DIR/"
if [ -d "$ROOT/public/assets" ]; then cp -r "$ROOT/public/assets" "$TEMP_DIR/"; fi
for img in "$ROOT/public"/*.jpg "$ROOT/public"/*.png "$ROOT/public"/*.gif "$ROOT/public"/*.jpeg; do
    [ -f "$img" ] && cp "$img" "$TEMP_DIR/"
done

# 离线包使用扁平布局：复制「扁平模式」启动脚本为根目录 start-server.py
cp "$ROOT/scripts/start-server-offline-flat.py" "$TEMP_DIR/start-server.py"

cat > "$TEMP_DIR/启动服务器.command" << 'EOF'
#!/bin/bash
cd "$(dirname "$0")"
echo -ne "\033]0;Word Add-in 服务器\007"
clear
if ! command -v python3 &> /dev/null; then
    echo "错误: 未找到 Python 3"; read -p "按回车退出..."; exit 1
fi
python3 start-server.py
read -p "按回车退出..."
EOF
chmod +x "$TEMP_DIR/启动服务器.command"

cat > "$TEMP_DIR/start-server.bat" << 'EOF'
@echo off
cd /d "%~dp0"
python --version >nul 2>&1 || (echo 未找到 Python & pause & exit /b 1)
python start-server.py
pause
EOF

cp "$ROOT/docs/USER_OFFLINE_INSTALL.md" "$TEMP_DIR/README.md"
cp "$ROOT/docs/故障排除指南.md" "$TEMP_DIR/" 2>/dev/null || true
cp "$ROOT/docs/快速诊断.txt" "$TEMP_DIR/" 2>/dev/null || true

chmod +x "$TEMP_DIR/start-server.py"

echo "创建 ZIP 包..."
(
    cd "$TEMP_DIR"
    zip -r "../$PACKAGE_NAME" . \
        -x "*.log" "node_modules/*" ".gitignore" "*.git/*"
)
rm -rf "$TEMP_DIR"
echo "✓ 离线安装包已创建: $PACKAGE_NAME"
