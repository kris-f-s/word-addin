# 插件部署指南

本指南说明如何将插件部署给非开发人员使用。

## 部署方案选择

### 方案一：部署到 Web 服务器（推荐）

将插件文件部署到一个稳定的 HTTPS 服务器上，用户只需要加载一次 manifest.xml 文件。

**优点：**
- 用户安装简单
- 更新方便（只需更新服务器文件）
- 不需要用户维护服务器

**缺点：**
- 需要有一个 HTTPS 服务器
- 需要有效的 SSL 证书

### 方案二：打包分发（适合内网）

将插件打包，用户需要在自己的服务器上部署。

**优点：**
- 可以内网部署
- 数据不经过外部服务器

**缺点：**
- 每个用户/组织需要自己部署
- 需要基本的服务器知识

## 方案一：部署到 Web 服务器

### 步骤 1: 准备服务器

你需要一个支持 HTTPS 的 Web 服务器，例如：
- 公司内部服务器
- 云服务器（AWS、Azure、阿里云等）
- GitHub Pages（需要自定义域名和 SSL）
- Netlify、Vercel 等静态托管服务

### 步骤 2: 上传文件

将以下文件上传到服务器：

```
word-addin/
├── manifest.xml          # 需要更新 URL
├── taskpane.html
├── taskpane.css
├── taskpane.js
├── commands.html
├── dialog.html
└── assets/               # 图标文件（如果有）
    ├── icon-16.png
    ├── icon-32.png
    ├── icon-64.png
    └── icon-80.png
```

**注意：** 不需要上传 `server.js`、`package.json` 等开发文件。

### 步骤 3: 更新 manifest.xml

修改 `manifest.xml` 中的所有 URL，将 `https://localhost:3000` 替换为你的服务器地址：

```xml
<!-- 示例：如果你的服务器是 https://yourcompany.com/word-addin -->
<SourceLocation DefaultValue="https://yourcompany.com/word-addin/taskpane.html"/>
<IconUrl DefaultValue="https://yourcompany.com/word-addin/assets/icon-32.png"/>
<!-- ... 其他 URL ... -->
```

### 步骤 4: 创建用户安装包

创建一个 ZIP 文件，只包含 `manifest.xml`：

```bash
cd word-addin
zip word-addin-install.zip manifest.xml
```

### 步骤 5: 提供安装说明

给用户提供简单的安装说明（见 `USER_INSTALL.md`）。

## 方案二：打包分发

### 步骤 1: 创建分发包

创建一个包含所有必要文件的 ZIP 包：

```bash
cd word-addin
zip -r word-addin-package.zip \
    manifest.xml \
    taskpane.html \
    taskpane.css \
    taskpane.js \
    commands.html \
    dialog.html \
    assets/ \
    -x "*.md" "*.sh" "server.js" "package.json" "node_modules/*" "*.pem" ".gitignore"
```

### 步骤 2: 提供部署说明

为用户提供部署说明（见 `USER_DEPLOY.md`）。

## 方案三：使用 Office Add-in 分发（企业级）

### 通过 SharePoint 分发

1. 将插件上传到 SharePoint App Catalog
2. 用户通过 **插入** > **加载项** > **获取加载项** 安装

### 通过中央部署

1. 管理员配置中央部署
2. 用户自动获得插件

## 更新插件

### 如果使用方案一（Web 服务器）

1. 更新服务器上的文件
2. 用户无需重新安装，下次打开 Word 时会自动使用新版本

### 如果使用方案二（打包分发）

1. 创建新版本的 ZIP 包
2. 通知用户下载并重新部署

## 注意事项

1. **SSL 证书**：必须使用有效的 SSL 证书，不能是自签名证书
2. **CORS**：确保服务器允许跨域请求（如果需要）
3. **文件路径**：确保所有文件的路径正确
4. **版本号**：更新插件时记得更新 `manifest.xml` 中的版本号

## 测试部署

部署后，在干净的 Word 环境中测试：

1. 卸载之前的插件版本
2. 使用新的 manifest.xml 安装
3. 测试所有功能

