$ErrorActionPreference = "Stop"
Set-Location $PSScriptRoot

python -m PyInstaller `
  --noconfirm `
  --clean `
  --onedir `
  --windowed `
  --name "PDF2Office" `
  --icon "assets\icon.ico" `
  --add-data "assets;assets" `
  --collect-all tkinterdnd2 `
  --collect-all docx `
  --collect-all pptx `
  app.py

Write-Host ""
Write-Host "构建完成: dist\PDF2Office\PDF2Office.exe"
Write-Host "请把整个 dist\PDF2Office 文件夹放在英文路径下使用（例如桌面）。"
