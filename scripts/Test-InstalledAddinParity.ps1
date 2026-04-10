$ErrorActionPreference = "Stop"

$root = Split-Path -Parent $PSScriptRoot
$package = Resolve-Path (Join-Path $root "package\THU-Formatter-Lite.dotm")
$startup = Join-Path $env:APPDATA "Microsoft\Word\STARTUP\THU-Formatter-Lite.dotm"

if (-not (Test-Path $startup)) {
  throw "Startup add-in missing: $startup"
}

$packageHash = Get-FileHash -LiteralPath $package.Path -Algorithm SHA256
$startupHash = Get-FileHash -LiteralPath $startup -Algorithm SHA256

if ($packageHash.Hash -ne $startupHash.Hash) {
  throw "Installed add-in does not match package build.`nPackage SHA256: $($packageHash.Hash)`nStartup SHA256: $($startupHash.Hash)"
}

Write-Host "Installed add-in matches package build."
Write-Host "Package SHA256: $($packageHash.Hash)"
Write-Host "Startup SHA256: $($startupHash.Hash)"
