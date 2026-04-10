$ErrorActionPreference = "Stop"

$root = Split-Path -Parent $PSScriptRoot
$workspaceRoot = Split-Path -Parent $root
$sampleDocx = Join-Path $workspaceRoot "thesis-format-engine\samples\mild_messy\sample_mild.docx"

if (-not (Test-Path $sampleDocx)) {
  throw "Missing sample DOCX: $sampleDocx"
}

$tempDirName = [string]::Concat(
  [char]0x8BBA, [char]0x6587, [char]0x6865, [char]0x63A5,
  [char]0x4E2D, [char]0x6587, [char]0x8DEF, [char]0x5F84
)
$docStem = [string]::Concat(
  [char]0x6865, [char]0x63A5, [char]0x4E2D, [char]0x6587,
  [char]0x8DEF, [char]0x5F84, [char]0x6837, [char]0x4F8B
)

$tempDir = Join-Path $env:TEMP $tempDirName
$docPath = Join-Path $tempDir ($docStem + ".docx")
$fixedDocx = Join-Path $tempDir ($docStem + "_fixed.docx")
$fixedPdf = Join-Path $tempDir ($docStem + "_fixed.pdf")
$reportJson = Join-Path $tempDir ($docStem + "_report.json")
$reportHtml = Join-Path $tempDir ($docStem + "_report.html")
$reportText = Join-Path $tempDir ($docStem + "_report.txt")
$logPath = Join-Path $tempDir ($docStem + "_fix_log.txt")

New-Item -ItemType Directory -Path $tempDir -Force | Out-Null
Copy-Item -LiteralPath $sampleDocx -Destination $docPath -Force

foreach ($artifact in @($fixedDocx, $fixedPdf, $reportJson, $reportHtml, $reportText, $logPath)) {
  if (Test-Path $artifact) {
    Remove-Item -LiteralPath $artifact -Force
  }
}

$word = $null
$document = $null

try {
  [Environment]::SetEnvironmentVariable("THU_ENGINE_SMOKE", "1", "Process")
  $word = New-Object -ComObject Word.Application
  $word.Visible = $false
  $word.DisplayAlerts = 0
  $document = $word.Documents.Open($docPath)
  $null = $word.Run("THU_Formatter_Addin.OneClickDetectAndFix")
  Start-Sleep -Seconds 2
}
finally {
  [Environment]::SetEnvironmentVariable("THU_ENGINE_SMOKE", $null, "Process")
  if ($null -ne $document) {
    $document.Close(0)
  }
  if ($null -ne $word) {
    $word.Quit()
  }
}

if (-not (Test-Path $logPath)) {
  throw "Missing bridge log: $logPath"
}

$logText = Get-Content -LiteralPath $logPath -Encoding UTF8 -Raw

if ($logText -notmatch "FLOW: entering CLI pipeline") {
  throw "CLI pipeline did not run."
}

if ($logText -match "cli_snapshot_savecopyas_failed=5892") {
  throw "Unexpected SaveCopyAs fallback remained in the log."
}

if ($logText -notmatch $docStem) {
  throw "UTF-8 log does not contain the expected Chinese path text."
}

Write-Host "Chinese-path bridge smoke passed."
Write-Host "Log: $logPath"
