const express = require('express');
const https = require('https');
const fs = require('fs');
const path = require('path');

const app = express();
const PORT = 3000;
const ROOT = __dirname;
const PUBLIC = path.join(ROOT, 'public');
const CERTS = path.join(ROOT, 'certs');

// 静态文件：URL 根路径对应 public/，与 manifest 中路径一致
app.use(express.static(PUBLIC));

app.get('/', (req, res) => {
    res.sendFile(path.join(PUBLIC, 'taskpane.html'));
});

const keyPath = path.join(CERTS, 'localhost-key.pem');
const certPath = path.join(CERTS, 'localhost.pem');

if (!fs.existsSync(keyPath) || !fs.existsSync(certPath)) {
    console.error('缺少 SSL 证书。请在项目根目录执行: bash scripts/generate-cert.sh');
    process.exit(1);
}

const options = {
    key: fs.readFileSync(keyPath).toString(),
    cert: fs.readFileSync(certPath).toString()
};

https.createServer(options, app).listen(PORT, () => {
    console.log(`Server running at https://localhost:${PORT}`);
    console.log('Make sure to trust the self-signed certificate in your browser.');
});
