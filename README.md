# AliCDN Batch Image Downloader

一个面向 Windows 的轻量级阿里 CDN 图片批量下载工具。把 Excel 中的一整列图片直链粘贴到界面，即可批量下载，无需逐条打开，也无需安装 Python 或 JDownloader。

> 本项目是独立开源工具，与 Alibaba.com、阿里巴巴集团或其关联公司无隶属、赞助或认可关系。

[English README](README_EN.md)

## 功能

- 一次粘贴几十、几百或更多图片链接
- 自动从 Excel 内容、Markdown 链接和混合文本中提取 HTTP/HTTPS URL
- 自动去除重复链接
- 保留 CDN 原始文件名，或按 `001、002、003...` 自动编号
- 下载失败自动重试，重试次数可调整
- 同名文件自动添加 `_2、_3`，不会覆盖已有图片
- 失败链接自动保存为 TXT 文件
- 自动记住下载路径、命名方式和重试次数
- 支持 JPG、JPEG、PNG、WebP、GIF、BMP、TIF、TIFF、AVIF
- 完全本地运行，不收集数据，不包含遥测

## 系统要求

- Windows 10 或 Windows 11
- Windows PowerShell 5.1（系统自带）
- 能够访问待下载图片的网络环境

## 使用方法

1. 从 [Releases](../../releases) 下载最新的 ZIP 文件。
2. 完整解压 ZIP，不要直接在压缩包内运行。
3. 双击 `双击启动.bat`。
4. 从 Excel 复制图片链接整列，点击“从剪贴板粘贴”。
5. 选择保存位置和文件命名方式。
6. 点击“一键开始下载”。

## 设置保存位置

程序会把个人设置保存在：

```text
%LOCALAPPDATA%\AlibabaImageDownloader\settings.json
```

该文件只保存在使用者自己的电脑中，不会随工具包共享，也不会上传到任何服务器。

## 常见问题

### 同事能否直接使用？

可以。把完整 ZIP 发给同事，对方解压后双击启动即可。每位使用者的下载路径和设置相互独立。

### Windows 阻止脚本运行怎么办？

启动文件仅为本次 PowerShell 进程使用 `ExecutionPolicy Bypass`，不会永久修改系统执行策略。如果公司电脑通过组织策略禁止 PowerShell 脚本，需要联系管理员放行。

### 为什么部分链接下载失败？

可能原因包括链接过期、CDN 防盗链、网络限制或目标文件已删除。程序会自动重试，并把失败链接保存为 TXT 文件。

## 隐私与安全

- 程序只访问使用者主动粘贴的图片 URL。
- 不上传链接、图片、使用记录或电脑信息。
- 不包含广告、统计代码、自动更新模块或第三方依赖。
- 源码为 PowerShell，可直接审查。

## 开发

主程序：`AlibabaImageDownloader.ps1`

启动器：`双击启动.bat`

项目使用 Windows Forms 构建图形界面，不需要额外构建步骤。

## 许可证

[MIT License](LICENSE)

