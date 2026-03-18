const express = require('express');
const https = require('https');
const fs = require('fs');
const path = require('path');

const app = express();
const PORT = 3000;

// 静态文件服务
app.use(express.static(__dirname));

// 根路径重定向到 taskpane.html
app.get('/', (req, res) => {
    res.sendFile(path.join(__dirname, 'taskpane.html'));
});

// 创建自签名证书（仅用于开发）
// 注意：在生产环境中，应该使用有效的 SSL 证书
const options = {
    key: fs.readFileSync(path.join(__dirname, 'localhost-key.pem')).toString(),
    cert: fs.readFileSync(path.join(__dirname, 'localhost.pem')).toString()
};

// 启动 HTTPS 服务器
https.createServer(options, app).listen(PORT, () => {
    console.log(`Server running at https://localhost:${PORT}`);
    console.log('Make sure to trust the self-signed certificate in your browser.');
});

