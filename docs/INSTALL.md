# 安装指南

## 快速开始

### 1. 安装依赖

```bash
cd word-addin
npm install
```

### 2. 生成 SSL 证书

```bash
bash scripts/generate-cert.sh
```

或者手动生成：

```bash
# 生成私钥
openssl genrsa -out localhost-key.pem 2048

# 生成证书
openssl req -new -x509 -key localhost-key.pem -out localhost.pem -days 365 \
    -subj "/C=CN/ST=State/L=City/O=Organization/CN=localhost"
```

### 3. 信任证书（macOS）

1. 打开 **Keychain Access**（钥匙串访问）
2. 将 `localhost.pem` 文件拖入左侧的 **login** 钥匙串
3. 双击导入的证书
4. 展开 **Trust** 部分
5. 将 **When using this certificate** 设置为 **Always Trust**
6. 关闭窗口，输入密码确认

### 4. 创建图标文件（可选）

插件需要以下图标文件，如果不存在，Word 会使用默认图标：

- `assets/icon-16.png` (16x16 像素)
- `assets/icon-32.png` (32x32 像素)
- `assets/icon-64.png` (64x64 像素)
- `assets/icon-80.png` (80x80 像素)

你可以创建简单的 PNG 图标，或者暂时跳过这一步（Word 会使用默认图标）。

### 5. 启动服务器

```bash
npm start
```

服务器将在 `https://localhost:3000` 启动。

### 6. 在 Word 中加载插件

#### 方法一：通过清单文件（推荐）

1. 打开 **Microsoft Word for Mac**
2. 转到 **插入** > **加载项** > **我的加载项**
3. 点击 **上传我的加载项**
4. 浏览并选择 `word-addin/manifest.xml` 文件
5. 如果出现安全警告，选择 **信任并加载**

#### 方法二：启用开发者模式

1. 打开终端，运行：
```bash
defaults write com.microsoft.Word OfficeWebAddinDeveloperExtras -bool true
```

2. 重启 Word

3. 在 Word 中，转到 **插入** > **加载项** > **我的加载项** > **上传我的加载项**

4. 选择 `manifest.xml` 文件

## 验证安装

1. 在 Word 中，你应该能看到 **开始** 选项卡中有一个新的按钮组
2. 点击 **显示文本格式化工具** 按钮，应该会打开任务窗格
3. 或者使用快捷键 `Shift+Alt+O+P`（需要任务窗格已打开或获得焦点）

## 故障排除

### 证书错误

如果浏览器或 Word 显示证书错误：

1. 确保已正确生成证书文件
2. 确保已在 Keychain Access 中信任证书
3. 尝试在浏览器中访问 `https://localhost:3000`，接受证书警告

### 插件无法加载

1. 检查服务器是否正在运行：`npm start`
2. 检查 `manifest.xml` 中的 URL 是否为 `https://localhost:3000`
3. 在浏览器中访问 `https://localhost:3000/taskpane.html` 查看是否有错误

### 快捷键不工作

Office Add-in 的快捷键功能有限。建议：

1. 使用 Word 功能区中的按钮打开任务窗格
2. 或者将任务窗格固定，使其始终可见

### 端口被占用

如果 3000 端口被占用，可修改根目录 `server.js` 或 `scripts/start-server.py` 中的 `PORT`，并同步更新 `manifest.xml` / `manifest-local.xml` 中的 URL。

## 下一步

安装完成后，请查看 `README.md` 了解如何使用插件。

