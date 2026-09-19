# AliCDN Batch Image Downloader

A lightweight Windows GUI for downloading image URLs in bulk. Paste a column of direct image links from Excel and download them without opening each URL or installing Python or JDownloader.

> This is an independent open-source project. It is not affiliated with, sponsored by, or endorsed by Alibaba.com, Alibaba Group, or their affiliates.

[中文说明](README.md)

## Features

- Paste dozens or hundreds of image URLs at once
- Extract HTTP/HTTPS URLs from Excel cells, Markdown links, and mixed text
- Remove duplicate URLs automatically
- Keep original CDN filenames or use sequential names such as `001`, `002`, `003`
- Retry failed downloads with a configurable retry count
- Avoid overwriting existing files by adding `_2`, `_3`, and so on
- Save failed URLs to a TXT report
- Remember the download folder, naming mode, and retry count
- Support JPG, JPEG, PNG, WebP, GIF, BMP, TIF, TIFF, and AVIF
- Run locally with no telemetry or data collection

## Requirements

- Windows 10 or Windows 11
- Windows PowerShell 5.1, included with Windows
- Network access to the image hosts being downloaded

## Usage

1. Download the latest ZIP from [Releases](../../releases).
2. Extract the complete archive. Do not run it from inside the ZIP file.
3. Double-click `双击启动.bat`.
4. Copy a column of image URLs from Excel and click the paste button.
5. Choose a destination folder and naming mode.
6. Click the blue download button.

## Local settings

Preferences are stored locally at:

```text
%LOCALAPPDATA%\AlibabaImageDownloader\settings.json
```

This file stays on each user's computer and is never included in the shared package or uploaded by the application.

## Privacy and security

- The application only requests URLs explicitly pasted by the user.
- It does not upload URLs, images, usage history, or device information.
- It contains no advertising, analytics, telemetry, or automatic updater.
- The PowerShell source is included and can be reviewed directly.

## Development

Main application: `AlibabaImageDownloader.ps1`

Launcher: `双击启动.bat`

The GUI uses Windows Forms and has no external build dependencies.

## License

[MIT License](LICENSE)

