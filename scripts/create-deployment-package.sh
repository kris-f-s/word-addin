#!/bin/bash
# 完整部署包：根目录扁平静态文件（与旧版一致，便于 nginx / 简单 Node）
set -e
ROOT="$(cd "$(dirname "$0")/.." && pwd)"
cd "$ROOT"

PACKAGE_NAME="word-addin-package.zip"
TEMP_DIR="word-addin-temp"

if [ -f "$PACKAGE_NAME" ]; then rm "$PACKAGE_NAME"; fi
if [ -d "$TEMP_DIR" ]; then rm -rf "$TEMP_DIR"; fi

mkdir -p "$TEMP_DIR"
echo "复制文件..."
cp manifest.xml "$TEMP_DIR/"
cp "$ROOT/public/"*.html "$TEMP_DIR/"
cp "$ROOT/public/taskpane.css" "$ROOT/public/taskpane.js" "$ROOT/public/logger.js" "$TEMP_DIR/"
if [ -d "$ROOT/public/assets" ]; then cp -r "$ROOT/public/assets" "$TEMP_DIR/"; fi
for img in "$ROOT/public"/*.jpg "$ROOT/public"/*.png "$ROOT/public"/*.gif "$ROOT/public"/*.jpeg; do
    [ -f "$img" ] && cp "$img" "$TEMP_DIR/"
done

cp "$ROOT/scripts/server-deployment-flat.js" "$TEMP_DIR/server.js"
if [ -f "$ROOT/package.json" ]; then cp "$ROOT/package.json" "$TEMP_DIR/"; fi

cat > "$TEMP_DIR/README.txt" << 'EOF'
Word 文本颜色格式化插件 - 部署包

静态部署：将本目录下 HTML/JS/CSS/assets 上传到 Web 服务器根目录，
保证 taskpane.html 的 URL 与 manifest.xml 中配置一致。

Node HTTPS：证书需为本目录下的 localhost.pem / localhost-key.pem（与旧版相同），
然后 npm install && npm start。
EOF

echo "创建 ZIP 包..."
(
    cd "$TEMP_DIR"
    zip -r "../$PACKAGE_NAME" . -x "*.md" "*.log" "node_modules/*" ".gitignore" "*.git/*"
)
rm -rf "$TEMP_DIR"
echo "✓ 部署包已创建: $PACKAGE_NAME"
