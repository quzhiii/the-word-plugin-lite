$ErrorActionPreference = "Stop"

$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$packageDir = Join-Path $root "package"
$dotm = Join-Path $packageDir "THU-Formatter-Lite.dotm"
$zipPath = Join-Path $packageDir "THU-Formatter-Lite.zip"
$tempDir = Join-Path $packageDir "_release_tmp"

if (-not (Test-Path $dotm)) {
  Write-Error "未找到 $dotm。请先完成 dotm 打包。"
}

& (Join-Path $root "VALIDATE_PACKAGE.ps1") | Out-Host

if (Test-Path $tempDir) {
  Remove-Item -Recurse -Force $tempDir
}
New-Item -ItemType Directory -Path $tempDir | Out-Null

Copy-Item $dotm (Join-Path $tempDir "THU-Formatter-Lite.dotm")
Copy-Item (Join-Path $root "INSTALL.ps1") (Join-Path $tempDir "INSTALL.ps1")
Copy-Item (Join-Path $root "UNINSTALL.ps1") (Join-Path $tempDir "UNINSTALL.ps1")
Copy-Item (Join-Path $root "README.zh-CN.md") (Join-Path $tempDir "README.zh-CN.md")

if (Test-Path $zipPath) {
  Remove-Item -Force $zipPath
}

Compress-Archive -Path (Join-Path $tempDir "*") -DestinationPath $zipPath -Force
Remove-Item -Recurse -Force $tempDir

Write-Host "发布包已生成: $zipPath"
