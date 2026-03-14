$ErrorActionPreference = "Stop"

$startup = Join-Path $env:APPDATA "Microsoft\Word\STARTUP"
$target = Join-Path $startup "THU-Formatter-Lite.dotm"

if (Test-Path $target) {
  Remove-Item -Path $target -Force
  Write-Host "已卸载: $target"
} else {
  Write-Host "未发现已安装模板: $target"
}

Write-Host "请重启 Word。"
