$ErrorActionPreference = "Stop"

Add-Type -AssemblyName System.IO.Compression.FileSystem

$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$packageFile = Join-Path $root "package\THU-Formatter-Lite.dotm"

if (-not (Test-Path $packageFile)) {
  Write-Error "未找到模板文件: $packageFile"
}

$zip = $null
try {
  $zip = [System.IO.Compression.ZipFile]::OpenRead($packageFile)
  $entryNames = $zip.Entries | ForEach-Object { $_.FullName }

  $required = @(
    "[Content_Types].xml",
    "word/vbaProject.bin",
    "customUI/customUI14.xml"
  )

  $missing = @()
  foreach ($item in $required) {
    if ($entryNames -notcontains $item) {
      $missing += $item
    }
  }

  if ($missing.Count -gt 0) {
    Write-Error ("dotm 包结构不完整，缺少: " + ($missing -join ", "))
  }

  Write-Host "包校验通过：$packageFile"
  Write-Host "已检测到 vbaProject.bin 与 customUI/customUI14.xml"
}
finally {
  if ($zip -ne $null) {
    $zip.Dispose()
  }
}
