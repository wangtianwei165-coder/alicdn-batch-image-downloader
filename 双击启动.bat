@echo off
cd /d "%~dp0"
powershell.exe -NoProfile -ExecutionPolicy Bypass -STA -File "%~dp0AlibabaImageDownloader.ps1"
if errorlevel 1 (
  echo.
  echo The tool could not start. Please take a screenshot of this window and send it back.
  pause
)

