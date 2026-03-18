#!/bin/bash
# 生成自签名 SSL 证书到 certs/（供 npm start / Python 服务器使用）

set -e
ROOT="$(cd "$(dirname "$0")/.." && pwd)"
mkdir -p "$ROOT/certs"
cd "$ROOT/certs"

echo "正在生成 SSL 证书（输出目录: certs/）..."

openssl genrsa -out localhost-key.pem 2048
openssl req -new -x509 -key localhost-key.pem -out localhost.pem -days 365 \
    -subj "/C=CN/ST=State/L=City/O=Organization/CN=localhost"

echo "证书生成完成！"
echo ""
echo "下一步："
echo "1. 在 macOS 的 Keychain Access 中导入 certs/localhost.pem"
echo "2. 双击证书，展开 Trust，设置为 Always Trust"
echo "3. 运行 npm start 或双击 启动服务器.command"
