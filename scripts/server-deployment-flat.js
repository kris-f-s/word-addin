/**
 * 扁平部署用 HTTPS 服务器（部署包内 HTML/JS 与 manifest 同级时使用）
 * 由 create-deployment-package.sh 复制为 zip 根目录下的 server.js
 */
const express = require('express');
const https = require('https');
const fs = require('fs');
const path = require('path');

const app = express();
const PORT = 3000;
const ROOT = __dirname;

app.use(express.static(ROOT));
app.get('/', (req, res) => {
    res.sendFile(path.join(ROOT, 'taskpane.html'));
});

const keyPath = path.join(ROOT, 'localhost-key.pem');
const certPath = path.join(ROOT, 'localhost.pem');
if (!fs.existsSync(keyPath) || !fs.existsSync(certPath)) {
    console.error('请在本目录放置 localhost-key.pem 与 localhost.pem（可用 openssl 生成）');
    process.exit(1);
}
https.createServer(
    { key: fs.readFileSync(keyPath), cert: fs.readFileSync(certPath) },
    app
).listen(PORT, () => {
    console.log(`https://localhost:${PORT}`);
});
