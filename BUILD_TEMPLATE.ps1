param(
  [string]$BaseDotmPath = "",
  [string]$OutputPath = ""
)

$ErrorActionPreference = "Stop"

Add-Type -AssemblyName System.IO.Compression.FileSystem

$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$packageDir = Join-Path $root "package"
$defaultBase = Join-Path $packageDir "THU-Formatter-Lite.base.dotm"
$defaultOutput = Join-Path $packageDir "THU-Formatter-Lite.dotm"
$customUiSource = Join-Path $root "customUI\customUI14.xml"
$tempDir = Join-Path $packageDir "_build_tmp"
$customUiTargetRelative = "customUI/customUI14.xml"
$customUiContentType = "application/vnd.ms-office.customUI+xml"
$customUiRelationshipType = "http://schemas.microsoft.com/office/2006/relationships/ui/extensibility"

if ([string]::IsNullOrWhiteSpace($BaseDotmPath)) {
  $BaseDotmPath = $defaultBase
}
if ([string]::IsNullOrWhiteSpace($OutputPath)) {
  $OutputPath = $defaultOutput
}

if (-not (Test-Path $BaseDotmPath)) {
  Write-Error "Base .dotm not found: $BaseDotmPath`nCreate one by copying Normal.dotm, importing src/vba/THU_Formatter_Addin.bas in Word, and saving it as THU-Formatter-Lite.base.dotm."
}
if (-not (Test-Path $customUiSource)) {
  Write-Error "Missing Ribbon XML source: $customUiSource"
}

$BaseDotmPath = (Resolve-Path $BaseDotmPath).Path
$outputDir = Split-Path -Parent $OutputPath
if (-not (Test-Path $outputDir)) {
  New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
}

function Ensure-CustomUiContentType {
  param([string]$ContentTypesPath)

  [xml]$xml = Get-Content -LiteralPath $ContentTypesPath
  $ns = New-Object System.Xml.XmlNamespaceManager($xml.NameTable)
  $ns.AddNamespace("ct", "http://schemas.openxmlformats.org/package/2006/content-types")
  $existing = $xml.SelectSingleNode("/ct:Types/ct:Override[@PartName='/$customUiTargetRelative']", $ns)
  $obsoleteNodes = @(
    $xml.SelectNodes("/ct:Types/ct:Override[@ContentType='$customUiContentType' and @PartName!='/$customUiTargetRelative']", $ns)
  )
  foreach ($node in $obsoleteNodes) {
    $xml.Types.RemoveChild($node) | Out-Null
  }
  if ($existing -eq $null) {
    $node = $xml.CreateElement("Override", $ns.LookupNamespace("ct"))
    $node.SetAttribute("PartName", "/$customUiTargetRelative")
    $node.SetAttribute("ContentType", $customUiContentType)
    $xml.Types.AppendChild($node) | Out-Null
  }
  $xml.Save($ContentTypesPath)
}

function Ensure-CustomUiRootRelationship {
  param([string]$RelsPath)

  [xml]$xml = Get-Content -LiteralPath $RelsPath
  $ns = New-Object System.Xml.XmlNamespaceManager($xml.NameTable)
  $ns.AddNamespace("rel", "http://schemas.openxmlformats.org/package/2006/relationships")
  $existing = $xml.SelectSingleNode("/rel:Relationships/rel:Relationship[@Type='$customUiRelationshipType']", $ns)
  if ($existing -eq $null) {
    $ids = @(
      $xml.SelectNodes("/rel:Relationships/rel:Relationship", $ns) |
      ForEach-Object { $_.Id }
    )
    $nextNumber = 1
    while ($ids -contains "rId$nextNumber") {
      $nextNumber += 1
    }

    $node = $xml.CreateElement("Relationship", $ns.LookupNamespace("rel"))
    $node.SetAttribute("Id", "rId$nextNumber")
    $node.SetAttribute("Type", $customUiRelationshipType)
    $node.SetAttribute("Target", $customUiTargetRelative)
    $xml.Relationships.AppendChild($node) | Out-Null
  } else {
    $existing.SetAttribute("Target", $customUiTargetRelative)
  }
  $xml.Save($RelsPath)
}

function New-ZipFromDirectory {
  param(
    [string]$SourceDir,
    [string]$DestinationPath
  )

  $zip = [System.IO.Compression.ZipFile]::Open($DestinationPath, [System.IO.Compression.ZipArchiveMode]::Create)
  try {
    $basePath = (Resolve-Path $SourceDir).Path
    $files = Get-ChildItem -LiteralPath $basePath -Recurse -File
    foreach ($file in $files) {
      $relativePath = $file.FullName.Substring($basePath.Length + 1).Replace("\", "/")
      [System.IO.Compression.ZipFileExtensions]::CreateEntryFromFile($zip, $file.FullName, $relativePath, [System.IO.Compression.CompressionLevel]::Optimal) | Out-Null
    }
  }
  finally {
    $zip.Dispose()
  }
}

if (Test-Path $tempDir) {
  Remove-Item -Recurse -Force $tempDir
}
New-Item -ItemType Directory -Path $tempDir | Out-Null

[System.IO.Compression.ZipFile]::ExtractToDirectory($BaseDotmPath, $tempDir)

$vbaProjectPath = Join-Path $tempDir "word\vbaProject.bin"
if (-not (Test-Path $vbaProjectPath)) {
  Remove-Item -Recurse -Force $tempDir
  Write-Error "Base .dotm does not contain word/vbaProject.bin: $BaseDotmPath"
}

$customUiDir = Join-Path $tempDir "customUI"
if (Test-Path $customUiDir) {
  Remove-Item -Recurse -Force $customUiDir
}
New-Item -ItemType Directory -Path $customUiDir | Out-Null
Copy-Item -LiteralPath $customUiSource -Destination (Join-Path $customUiDir "customUI14.xml") -Force

Ensure-CustomUiContentType -ContentTypesPath (Join-Path $tempDir "[Content_Types].xml")
Ensure-CustomUiRootRelationship -RelsPath (Join-Path $tempDir "_rels\.rels")

if (Test-Path $OutputPath) {
  Remove-Item -LiteralPath $OutputPath -Force
}
New-ZipFromDirectory -SourceDir $tempDir -DestinationPath $OutputPath
Remove-Item -Recurse -Force $tempDir

$zip = $null
try {
  $zip = [System.IO.Compression.ZipFile]::OpenRead($OutputPath)
  $entryNames = $zip.Entries | ForEach-Object { $_.FullName.Replace("\", "/") }
  foreach ($required in @("[Content_Types].xml", "word/vbaProject.bin", $customUiTargetRelative)) {
    if ($entryNames -notcontains $required) {
      Write-Error "Built template is missing required entry: $required"
    }
  }
}
finally {
  if ($zip -ne $null) {
    $zip.Dispose()
  }
}

Write-Host "Template built: $OutputPath"
Write-Host "Base template: $BaseDotmPath"
Write-Host "Ribbon XML injected from: $customUiSource"
Write-Host "Next: run .\VALIDATE_PACKAGE.ps1 and .\PACK_RELEASE.ps1"