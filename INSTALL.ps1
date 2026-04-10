param(
  [switch]$ForceCloseWord,
  [switch]$SkipCliBridge
)

$ErrorActionPreference = "Stop"

$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$packageFile = Join-Path $root "package\THU-Formatter-Lite.dotm"
$validator = Join-Path $root "VALIDATE_PACKAGE.ps1"
$engineCommand = $null

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

$wordProcesses = @(Get-Process WINWORD -ErrorAction SilentlyContinue)
if ($wordProcesses.Count -gt 0) {
  if ($ForceCloseWord) {
    Write-Host "Closing running Word processes..."
    $wordProcesses | Stop-Process -Force
    Start-Sleep -Seconds 2
  } else {
    Write-Error "Word is still running and has locked $target`nClose all Word windows first, then rerun INSTALL.ps1.`nIf needed, use: powershell -ExecutionPolicy Bypass -File .\INSTALL.ps1 -ForceCloseWord"
  }
}

Copy-Item -Path $packageFile -Destination $target -Force

$packageHash = Get-FileHash -LiteralPath $packageFile -Algorithm SHA256
$targetHash = Get-FileHash -LiteralPath $target -Algorithm SHA256
if ($packageHash.Hash -ne $targetHash.Hash) {
  throw "Installed add-in hash mismatch.`nPackage SHA256: $($packageHash.Hash)`nStartup SHA256: $($targetHash.Hash)"
}

if (-not $SkipCliBridge) {
  $engineCommand = Get-Command thesis-engine -ErrorAction SilentlyContinue
  if ($null -ne $engineCommand) {
    [Environment]::SetEnvironmentVariable("THU_ENGINE_MODE", "cli", "User")
    [Environment]::SetEnvironmentVariable("THU_ENGINE_CMD", $engineCommand.Source, "User")
    [Environment]::SetEnvironmentVariable("THU_ENGINE_PROFILE", "tsinghua-thesis", "User")
    [Environment]::SetEnvironmentVariable("THU_ENGINE_FIX_MODE", "full", "User")
  }
}

Write-Host "Install complete."
Write-Host "Startup: $startup"
Write-Host "Template: $target"
Write-Host "Package SHA256: $($packageHash.Hash)"
Write-Host "Startup SHA256: $($targetHash.Hash)"
Write-Host "Restart Word and check whether the THU Formatter tab appears."
Write-Host ""
if ($null -ne $engineCommand) {
  Write-Host "CLI bridge configured."
  Write-Host "THU_ENGINE_MODE=cli"
  Write-Host "THU_ENGINE_CMD=$($engineCommand.Source)"
  Write-Host "THU_ENGINE_PROFILE=tsinghua-thesis"
  Write-Host "THU_ENGINE_FIX_MODE=full"
} elseif (-not $SkipCliBridge) {
  Write-Host "thesis-engine not found on PATH, so the add-in will fall back to legacy VBA mode."
  Write-Host "If needed, install thesis-engine first, then rerun INSTALL.ps1."
}

Write-Host ""
Write-Host "Recommended verification:"
Write-Host "  powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\Test-InstalledAddinParity.ps1"
Write-Host "  powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\Test-StartupCompileSmoke.ps1"
