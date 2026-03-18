# 创建图标文件

插件需要以下图标文件。如果这些文件不存在，Word 会使用默认图标，但建议创建自定义图标以获得更好的用户体验。

## 需要的图标尺寸

- `assets/icon-16.png` - 16x16 像素
- `assets/icon-32.png` - 32x32 像素  
- `assets/icon-64.png` - 64x64 像素
- `assets/icon-80.png` - 80x80 像素

## 创建图标的方法

### 方法一：使用在线工具

1. 访问 [Favicon Generator](https://www.favicon-generator.org/) 或类似工具
2. 上传一个较大的图标（至少 512x512 像素）
3. 下载生成的不同尺寸图标
4. 将图标重命名并放入 `assets/` 目录

### 方法二：使用图像编辑软件

1. 使用 Photoshop、GIMP 或其他图像编辑软件
2. 创建一个简单的图标设计（例如：文本格式化工具的图标）
3. 导出为不同尺寸的 PNG 文件
4. 确保背景透明（如果需要）

### 方法三：使用命令行工具（ImageMagick）

如果你安装了 ImageMagick：

```bash
# 创建 assets 目录
mkdir -p assets

# 从一个大图标生成不同尺寸（假设你有一个 512x512 的 source.png）
convert source.png -resize 16x16 assets/icon-16.png
convert source.png -resize 32x32 assets/icon-32.png
convert source.png -resize 64x64 assets/icon-64.png
convert source.png -resize 80x80 assets/icon-80.png
```

### 方法四：使用占位图标

如果暂时不需要自定义图标，可以创建简单的占位图标：

```bash
mkdir -p assets

# 使用 ImageMagick 创建简单的彩色方块作为占位符
convert -size 16x16 xc:#0078d4 assets/icon-16.png
convert -size 32x32 xc:#0078d4 assets/icon-32.png
convert -size 64x64 xc:#0078d4 assets/icon-64.png
convert -size 80x80 xc:#0078d4 assets/icon-80.png
```

## 图标设计建议

- 使用简洁、清晰的设计
- 确保在小尺寸下仍然清晰可见
- 使用与 Office 风格一致的颜色（如蓝色 #0078d4）
- 考虑使用文本格式化相关的图标元素（如字母 "A" 或颜色调色板）

## 注意事项

- 图标文件必须是 PNG 格式
- 建议使用透明背景
- 确保图标在不同背景下都清晰可见

