# 插件分发指南（管理员版）

本指南帮助管理员将插件分发给非开发人员使用。

## 分发方案对比

| 方案 | 适用场景 | 用户难度 | 维护成本 |
|------|---------|---------|---------|
| **Web 服务器部署** | 公司统一部署 | ⭐ 简单 | ⭐ 低 |
| **打包分发** | 内网/多组织 | ⭐⭐ 中等 | ⭐⭐ 中等 |
| **Office Store** | 公开分发 | ⭐ 简单 | ⭐ 低（需审核） |

## 方案一：Web 服务器部署（推荐）

### 适用场景
- 公司内部统一使用
- 有 IT 部门维护服务器
- 希望统一更新和管理

### 部署步骤

#### 1. 准备服务器

选择一个 HTTPS Web 服务器：
- 公司内部服务器
- 云服务器（AWS、Azure、阿里云等）
- 静态托管服务（需要自定义域名）

#### 2. 更新 manifest.xml

使用提供的脚本更新 URL：

```bash
node scripts/update-manifest.js https://yourcompany.com/word-addin
```

或手动编辑 `manifest.xml`，将所有 `https://localhost:3000` 替换为你的服务器地址。

#### 3. 上传文件到服务器

上传以下文件到服务器的 Web 目录：

```
word-addin/
├── manifest.xml          # 已更新 URL
├── taskpane.html
├── taskpane.css
├── taskpane.js
├── commands.html
├── dialog.html
└── assets/              # 图标文件
    ├── icon-16.png
    ├── icon-32.png
    ├── icon-64.png
    └── icon-80.png
```

#### 4. 配置 SSL 证书

确保服务器使用有效的 SSL 证书（不能是自签名证书）。

#### 5. 测试部署

在浏览器中访问：
- `https://yourcompany.com/word-addin/taskpane.html`

应该能看到插件界面。

#### 6. 创建用户安装包

```bash
bash scripts/create-package.sh
```

这会创建 `word-addin-install.zip`，只包含 `manifest.xml`。

#### 7. 分发给用户

将以下内容分发给用户：
- `word-addin-install.zip` 安装包
- `USER_INSTALL.md` 用户安装指南

### 用户安装步骤

用户只需要：
1. 解压 `word-addin-install.zip`
2. 在 Word 中加载 `manifest.xml`
3. 完成！

### 更新插件

1. 更新服务器上的文件
2. 用户无需重新安装，自动使用新版本

## 方案二：打包分发

### 适用场景
- 多个组织/部门独立部署
- 内网环境，无法访问外网
- 需要数据本地化

### 部署步骤

#### 1. 创建部署包

```bash
bash scripts/create-deployment-package.sh
```

这会创建 `word-addin-package.zip`，包含所有必要文件。

#### 2. 分发给用户

将以下内容分发给用户：
- `word-addin-package.zip` 部署包
- `USER_DEPLOY.md` 部署指南

### 用户部署步骤

用户需要：
1. 解压部署包
2. 部署到自己的服务器（需要基本 IT 知识）
3. 在 Word 中加载 manifest.xml

详见 `USER_DEPLOY.md`。

## 方案三：Office Store（可选）

### 适用场景
- 公开分发
- 希望用户通过 Office Store 安装

### 步骤

1. 准备插件包
2. 提交到 Office Store 审核
3. 审核通过后，用户可以通过 **插入** > **加载项** > **获取加载项** 安装

**注意：** 需要有效的 SSL 证书和完整的隐私政策。

## 推荐方案

### 对于大多数公司：方案一（Web 服务器部署）

**优点：**
- ✅ 用户安装最简单
- ✅ 更新方便
- ✅ 统一管理
- ✅ 无需用户维护服务器

**实施步骤：**
1. IT 部门部署到公司服务器
2. 创建安装包分发给用户
3. 用户只需加载 manifest.xml

### 对于多组织/内网：方案二（打包分发）

**优点：**
- ✅ 数据本地化
- ✅ 各组织独立部署
- ✅ 不依赖外网

**缺点：**
- ❌ 用户需要基本 IT 知识
- ❌ 每个组织需要自己维护

## 分发清单

### 方案一（Web 服务器）分发内容：

- [ ] `word-addin-install.zip` - 用户安装包
- [ ] `USER_INSTALL.md` - 用户安装指南
- [ ] 服务器 URL 说明

### 方案二（打包分发）分发内容：

- [ ] `word-addin-package.zip` - 完整部署包
- [ ] `USER_DEPLOY.md` - 用户部署指南
- [ ] `README.md` - 项目说明

## 支持用户

### 常见问题

**Q: 用户说插件无法加载？**
- 检查服务器是否正常运行
- 检查 URL 是否正确
- 检查 SSL 证书是否有效

**Q: 用户说格式没有应用？**
- 这是使用问题，参考 `USER_INSTALL.md` 中的常见问题部分

**Q: 如何更新插件？**
- 方案一：更新服务器文件即可
- 方案二：提供新版本部署包

### 联系方式

为用户提供支持渠道：
- IT 支持邮箱
- 内部文档链接
- 帮助文档

## 安全注意事项

1. **SSL 证书**：必须使用有效证书，不能是自签名
2. **服务器安全**：确保服务器安全配置
3. **访问控制**：如果需要，可以配置访问控制
4. **更新机制**：建立安全的更新机制

## 版本管理

建议使用版本号管理：

1. 更新 `manifest.xml` 中的版本号
2. 记录更新日志
3. 通知用户更新（方案二）

## 总结

**最简单的方式（推荐）：**
1. 部署到 Web 服务器
2. 使用 `bash scripts/create-package.sh` 创建安装包
3. 分发 `word-addin-install.zip` 和 `USER_INSTALL.md`
4. 用户只需加载 manifest.xml，完成！

