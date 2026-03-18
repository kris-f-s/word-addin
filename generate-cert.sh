#!/bin/bash

# 生成自签名 SSL 证书用于本地开发

echo "正在生成 SSL 证书..."

# 生成私钥
openssl genrsa -out localhost-key.pem 2048

# 生成证书
openssl req -new -x509 -key localhost-key.pem -out localhost.pem -days 365 \
    -subj "/C=CN/ST=State/L=City/O=Organization/CN=localhost"

echo "证书生成完成！"
echo ""
echo "下一步："
echo "1. 在 macOS 的 Keychain Access 中导入 localhost.pem"
echo "2. 双击证书，展开 Trust，设置为 'Always Trust'"
echo "3. 运行 npm start 启动服务器"

