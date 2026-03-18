# 离线部署指南（不使用外部服务器）

本指南说明如何将插件部署为完全离线版本，用户可以在本地使用，无需连接外部服务器。

## 核心思路

Office Add-in 必须使用 HTTPS 协议，不能直接使用本地文件。解决方案是：
- 在用户本地运行一个简单的 HTTP 服务器
- 服务器提供 HTTPS 服务（使用自签名证书）
- 插件通过 `https://localhost:3000` 访问

## 方案对比

| 方案 | 优点 | 缺点 | 适用场景 |
|------|------|------|---------|
| **Python HTTP 服务器** | macOS 自带，无需安装 | Windows 需要安装 Python | macOS 用户（推荐） |
| **Node.js 服务器** | 功能完整 | 需要安装 Node.js | 已有 Node.js 环境 |
| **打包可执行文件** | 最简单，双击运行 | 需要编译打包 | 最佳用户体验（未来） |

## 方案一：Python HTTP 服务器（推荐）

### 适用场景
- **macOS 用户**：系统自带 Python 3，无需安装
- **Windows 用户**：需要安装 Python 3（[下载地址](https://www.python.org/downloads/)）

### 实现原理

1. 使用 Python 的 `http.server` 模块提供 HTTP 服务
2. 使用 `ssl` 模块提供 HTTPS 支持
3. 自动生成自签名 SSL 证书（首次运行）
4. 提供简单的启动脚本，用户双击即可运行

### 管理员操作步骤

#### 1. 创建离线安装包

```bash
./create-offline-package.sh
```

这会创建 `word-addin-offline.zip`，包含：
- 所有插件文件
- Python 服务器脚本 (`start-server.py`)
- macOS 启动脚本 (`启动服务器.command`)
- Windows 启动脚本 (`start-server.bat`)
- 本地 manifest 文件 (`manifest-local.xml`)
- 用户安装指南

#### 2. 分发给用户

将以下内容分发给用户：
- `word-addin-offline.zip` - 离线安装包
- `USER_OFFLINE_INSTALL.md` - 用户安装指南（可选，已包含在包中）

### 用户使用步骤

1. **解压插件包**到任意位置
2. **启动本地服务器**：
   - macOS: 双击 `启动服务器.command`
   - Windows: 双击 `start-server.bat`
3. **在 Word 中加载插件**：
   - 插入 > 加载项 > 我的加载项 > 上传我的加载项
   - 选择 `manifest-local.xml`
4. **使用插件**

### 技术细节

- **端口**: 默认使用 3000 端口
- **证书**: 自动生成自签名证书（首次运行）
- **HTTPS**: 使用 SSL/TLS 加密
- **CORS**: 已配置跨域支持

## 方案二：Node.js 服务器

如果用户已有 Node.js 环境，可以使用原有的 `server.js`。

### 用户使用步骤

1. 解压插件包
2. 运行 `npm install`（首次使用）
3. 运行 `npm start`
4. 在 Word 中加载 `manifest-local.xml`

## 方案三：打包为可执行文件（未来优化）

### 使用 pkg 打包 Node.js 服务器

```bash
npm install -g pkg
pkg server.js --targets node18-macos-x64,node18-win-x64
```

这会生成可执行文件，用户无需安装 Node.js。

## 文件说明

### 核心文件

- `start-server.py` - Python HTTP 服务器（主要方案）
- `启动服务器.command` - macOS 启动脚本
- `start-server.bat` - Windows 启动脚本
- `manifest-local.xml` - 本地使用的 manifest 文件（使用 localhost:3000）

### 用户文档

- `USER_OFFLINE_INSTALL.md` - 用户安装和使用指南

## 常见问题

### Q: 为什么需要本地服务器？

A: Office Add-in 要求使用 HTTPS 协议，不能直接使用 `file://` 协议。本地服务器提供 HTTPS 服务。

### Q: 安全吗？

A: 是的。使用的是自签名证书，仅用于本地通信，不会连接到外部服务器。所有数据都在本地处理。

### Q: 每次都要启动服务器吗？

A: 是的。每次使用插件前需要启动本地服务器。可以设置开机自动启动（需要一些技术知识）。

### Q: 可以修改端口吗？

A: 可以。修改 `start-server.py` 中的 `PORT = 3000`，同时修改 `manifest-local.xml` 中所有 `localhost:3000` 为新端口。

## 优势

✅ **完全离线**：不需要外部服务器  
✅ **简单易用**：用户只需双击启动脚本  
✅ **跨平台**：支持 macOS 和 Windows  
✅ **自动配置**：自动生成 SSL 证书  
✅ **零依赖**：macOS 用户无需安装任何软件  

## 总结

**推荐方案：Python HTTP 服务器**

1. 管理员运行 `./create-offline-package.sh` 创建安装包
2. 用户解压后双击启动脚本
3. 在 Word 中加载 manifest-local.xml
4. 完成！

这是最简单、最实用的离线部署方案。

