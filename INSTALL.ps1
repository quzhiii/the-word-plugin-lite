$ErrorActionPreference = "Stop"

$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$packageFile = Join-Path $root "package\THU-Formatter-Lite.dotm"
$validator = Join-Path $root "VALIDATE_PACKAGE.ps1"

if (-not (Test-Path $packageFile)) {
  Write-Error "Missing template file: $packageFile`nRun BUILD_TEMPLATE.ps1 first so package\THU-Formatter-Lite.dotm exists."
}

if (Test-Path $validator) {
  & $validator | Out-Host
}

try {
  Unblock-File -Path $packageFile -ErrorAction SilentlyContinue
} catch {
}

$startup = Join-Path $env:APPDATA "Microsoft\Word\STARTUP"
if (-not (Test-Path $startup)) {
  New-Item -ItemType Directory -Path $startup -Force | Out-Null
}

$target = Join-Path $startup "THU-Formatter-Lite.dotm"
Copy-Item -Path $packageFile -Destination $target -Force

Write-Host "Install complete."
Write-Host "Startup: $startup"
Write-Host "Template: $target"
Write-Host "Restart Word and check whether the THU Formatter tab appears."
Write-Host ""
Write-Host "Optional: configure thesis-format-engine bridge:"
Write-Host "  setx THU_ENGINE_MODE cli"
Write-Host "  setx THU_ENGINE_CMD `"thesis-engine`""
Write-Host "  setx THU_ENGINE_PROFILE `"tsinghua-thesis`""
Write-Host "  setx THU_ENGINE_FIX_MODE `"full`"    # optional: enable medium-risk table normalization"