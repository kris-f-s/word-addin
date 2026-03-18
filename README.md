# Word 文本颜色格式化插件

这是一个 Microsoft Word for Mac 的 Office Add-in 插件，可以通过快捷键或按钮打开一个任务窗格，用于批量格式化文档中的文本颜色。

## 功能特性

- **快捷键支持**: 使用 `Shift+Alt+O+P` 快速打开任务窗格
- **多规则支持**: 可以添加多条格式化规则
- **文本搜索**: 根据输入的文本搜索文档中的匹配内容
- **颜色格式化**: 同时设置字体颜色和背景颜色
- **动态添加规则**: 点击加号按钮可以添加新的规则行

## 项目结构

```
word-addin/
├── manifest.xml          # Office Add-in 清单文件
├── taskpane.html         # 主界面 HTML
├── taskpane.css          # 样式文件
├── taskpane.js           # 主要逻辑代码
├── commands.html         # 命令处理文件
├── server.js             # 本地开发服务器
├── package.json          # Node.js 依赖配置
└── README.md             # 说明文档
```

## 安装步骤

### 1. 安装 Node.js 依赖

```bash
cd word-addin
npm install
```

### 2. 生成 SSL 证书（用于本地 HTTPS 服务器）

由于 Office Add-in 要求使用 HTTPS，需要生成自签名证书：

```bash
# 生成私钥
openssl genrsa -out localhost-key.pem 2048

# 生成证书
openssl req -new -x509 -key localhost-key.pem -out localhost.pem -days 365 -subj "/CN=localhost"
```

### 3. 启动本地服务器

```bash
npm start
```

服务器将在 `https://localhost:3000` 启动。

### 4. 信任自签名证书（macOS）

1. 打开 Keychain Access（钥匙串访问）
2. 将 `localhost.pem` 拖入 "login" 钥匙串
3. 双击证书，展开 "Trust"，将 "When using this certificate" 设置为 "Always Trust"

### 5. 在 Word 中加载插件

#### 方法一：通过清单文件加载（推荐）

1. 打开 Microsoft Word for Mac
2. 转到 **插入** > **加载项** > **我的加载项**
3. 点击 **上传我的加载项**
4. 选择 `manifest.xml` 文件
5. 如果提示安全警告，选择 **信任并加载**

#### 方法二：通过侧加载（开发模式）

1. 打开终端，运行：
```bash
defaults write com.microsoft.Word OfficeWebAddinDeveloperExtras -bool true
```

2. 重启 Word

3. 在 Word 中，转到 **插入** > **加载项** > **我的加载项** > **上传我的加载项**

4. 选择 `manifest.xml` 文件

## 使用方法

### 打开任务窗格

- **方法一**: 使用快捷键 `Shift+Alt+O+P`
- **方法二**: 在 Word 的 **开始** 选项卡中，点击 **显示文本格式化工具** 按钮

### 添加格式化规则

1. 在第一个输入框中输入要搜索的文本
2. 在第二个颜色选择器中选择字体颜色
3. 在第三个颜色选择器中选择背景颜色
4. 点击 **+** 按钮可以添加更多规则

### 应用格式化

1. 配置完所有规则后，点击 **确认** 按钮
2. 插件会在文档中搜索所有匹配的文本
3. 将匹配的文本应用相应的颜色格式

### 取消操作

点击 **取消** 按钮可以清空所有规则并重新开始。

## 技术说明

### Office.js API

本插件使用 Office.js API 与 Word 文档交互：

- `Word.run()`: 执行 Word API 操作
- `body.search()`: 搜索文档中的文本
- `range.font.color`: 设置字体颜色
- `range.font.highlightColor`: 设置背景颜色

### 快捷键实现

快捷键 `Shift+Alt+O+P` 通过 JavaScript 事件监听实现。由于 Office Add-in 的限制，快捷键需要在任务窗格获得焦点时才能工作。

### 颜色格式

- **字体颜色**: 使用 `font.color` 属性
- **背景颜色**: 使用 `font.highlightColor` 属性（Word 的高亮功能）

## 开发说明

### 修改清单文件

如果需要修改插件的 ID、名称或其他配置，编辑 `manifest.xml` 文件。

**重要**: 修改后需要重新加载插件。

### 调试

1. 在浏览器中打开 `https://localhost:3000/taskpane.html` 查看界面
2. 在 Word 中，使用 **工具** > **脚本编辑器** 查看控制台日志
3. 使用 `console.log()` 输出调试信息

### 常见问题

#### 1. 证书错误

如果遇到证书错误，确保：
- 已正确生成 SSL 证书
- 已在 Keychain Access 中信任证书
- 服务器正在运行

#### 2. 插件无法加载

- 检查 `manifest.xml` 中的 URL 是否正确
- 确保服务器正在运行
- 检查浏览器控制台是否有错误

#### 3. 快捷键不工作

- 确保任务窗格已获得焦点
- 尝试使用按钮方式打开任务窗格

## 生产环境部署

在生产环境中，需要：

1. 使用有效的 SSL 证书（不是自签名证书）
2. 将文件部署到 HTTPS 服务器
3. 更新 `manifest.xml` 中的 URL
4. 在 Office Store 或通过企业部署分发插件

## 许可证

MIT License

