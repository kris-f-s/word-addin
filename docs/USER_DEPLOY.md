# 用户部署指南（本地服务器）

本指南适用于需要在本地服务器上部署插件的用户。

## 前提条件

- 一台可以运行 Web 服务器的计算机（Windows、macOS 或 Linux）
- Node.js（版本 12 或更高）
- 有效的 SSL 证书（生产环境）或自签名证书（测试环境）
- 基本的命令行操作知识

## 部署步骤

### 步骤 1: 获取插件包

从管理员获取 `word-addin-package.zip` 文件并解压。

### 步骤 2: 安装 Node.js

如果还没有安装 Node.js：

1. 访问 [Node.js 官网](https://nodejs.org/)
2. 下载并安装 LTS 版本
3. 打开终端/命令提示符，验证安装：
   ```bash
   node --version
   npm --version
   ```

### 步骤 3: 安装依赖

在插件目录中运行：

```bash
cd word-addin-package
npm install express
```

### 步骤 4: 配置服务器

#### 选项 A: 使用提供的 server.js（开发/测试）

- **部署包 zip（扁平目录）**：`server.js` 与 `taskpane.html` 同级；在该目录用 `openssl` 生成 `localhost.pem` / `localhost-key.pem`，或从其他环境复制证书到同级目录。
- **完整开发仓库**：在项目根目录执行 `bash scripts/generate-cert.sh`（证书写入 `certs/`），再 `npm start`；静态文件由 `public/` 提供。

然后按需修改 `server.js` 中的 `PORT`，并更新 `manifest.xml` 中的 URL。

#### 选项 B: 使用其他 Web 服务器

如果你有 Apache、Nginx 或其他 Web 服务器：

1. 将插件文件复制到服务器的 Web 根目录
2. 配置 HTTPS
3. 更新 `manifest.xml` 中的 URL

### 步骤 5: 启动服务器

#### 使用 Node.js 服务器：

```bash
node server.js
```

服务器将在 `https://localhost:3000` 启动（或你配置的端口）。

#### 使用其他服务器：

按照你的服务器文档配置并启动。

### 步骤 6: 测试服务器

在浏览器中访问：
- `https://your-server-address/taskpane.html`

如果能看到界面，说明服务器配置正确。

### 步骤 7: 在 Word 中安装

1. 打开 Microsoft Word
2. **插入** > **加载项** > **我的加载项**
3. 点击 **上传我的加载项**
4. 选择 `manifest.xml` 文件
5. 如果出现安全提示，选择 **信任并加载**

## 配置说明

### 更新 manifest.xml 中的 URL

打开 `manifest.xml`，将所有 `https://localhost:3000` 替换为你的实际服务器地址：

```xml
<!-- 示例 -->
<SourceLocation DefaultValue="https://your-server.com/word-addin/taskpane.html"/>
<IconUrl DefaultValue="https://your-server.com/word-addin/assets/icon-32.png"/>
```

### SSL 证书

**生产环境：**
- 必须使用有效的 SSL 证书（从证书颁发机构获取）
- 不能使用自签名证书

**测试环境：**
- 可以使用自签名证书
- 需要在浏览器和 Word 中信任证书

### 防火墙配置

确保服务器端口（如 3000）在防火墙中开放。

## 维护

### 更新插件

1. 停止服务器
2. 替换插件文件（保留配置文件）
3. 重启服务器

### 查看日志

如果使用 Node.js 服务器，日志会输出到控制台。

### 自动启动（可选）

可以配置服务器在系统启动时自动运行：

**macOS (使用 launchd):**
创建 `~/Library/LaunchAgents/com.wordaddin.server.plist`

**Windows (使用任务计划程序):**
创建计划任务运行 `node server.js`

**Linux (使用 systemd):**
创建 systemd service 文件

## 故障排除

### 服务器无法启动

1. 检查端口是否被占用
2. 检查 SSL 证书文件是否存在
3. 查看错误消息

### Word 无法连接

1. 检查服务器是否正在运行
2. 检查 URL 是否正确
3. 检查防火墙设置
4. 检查 SSL 证书是否有效

### 证书错误

1. 确保使用有效的 SSL 证书
2. 在浏览器中测试访问，接受证书警告
3. 检查证书是否过期

## 获取帮助

如果遇到问题：
1. 查看服务器日志
2. 检查网络连接
3. 联系 IT 支持或插件管理员

