# Word 文本颜色格式化插件

Microsoft Word 的 Office 加载项：通过任务窗格按规则批量设置文本颜色与高亮。

## 目录结构

```
word-addin/
├── public/                 # 插件前端（对应 https://localhost:3000/ 根路径）
│   ├── taskpane.html/js/css
│   ├── commands.html, dialog.html, logger.js, test-server.html
│   └── assets/             # 图标等
├── certs/                  # 本地 HTTPS 证书（gitignore，首次用脚本生成）
├── scripts/                # 打包、证书、Python 服务器、部署用 Node 脚本
├── docs/                   # 安装、部署、排错等文档
├── manifest.xml            # 清单（生产 / 自定义域名）
├── manifest-local.xml      # 本地开发清单
├── server.js               # 开发用 Node HTTPS 服务器
├── package.json
├── 启动服务器.command       # macOS：双击启动 Python 服务器
└── start-server.bat        # Windows：同上
```

**说明**：`manifest.xml` 中的 URL 仍为 `https://localhost:3000/taskpane.html` 等，与重组前一致；Node/Python 服务器会把 `public/` 映射到网站根路径。

## 快速开始（开发）

```bash
cd word-addin
npm install
bash scripts/generate-cert.sh   # 在 certs/ 生成 localhost.pem
npm start                       # https://localhost:3000
```

或在 macOS 双击 **`启动服务器.command`**（Python，无需 Node）。

证书信任、在 Word 中上传 `manifest.xml` 等步骤见 **`docs/INSTALL.md`**。

## 常用脚本（均在项目根执行）

| 操作 | 命令 |
|------|------|
| 用户安装包（仅 manifest） | `bash scripts/create-package.sh` |
| 完整部署 ZIP（扁平静态 + 扁平 Node） | `bash scripts/create-deployment-package.sh` |
| 离线包（最终用户扁平目录） | `bash scripts/create-offline-package.sh` |
| 批量改 manifest 里的域名 | `node scripts/update-manifest.js https://你的域名/word-addin` |

更多说明见 **`docs/`** 下各主题文档。

## 功能概要

- 快捷键 / 功能区按钮打开任务窗格  
- 多规则：搜索文本 + 字体色 + 高亮色  
- 详见原 **`docs/快捷键说明.md`** 等  

## 许可证

MIT License
