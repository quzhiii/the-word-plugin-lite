$ErrorActionPreference = "Stop"

Add-Type -AssemblyName System.IO.Compression.FileSystem

$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$packageFile = Join-Path $root "package\THU-Formatter-Lite.dotm"

if (-not (Test-Path $packageFile)) {
  Write-Error "Missing template file: $packageFile"
}

$zip = $null
try {
  $zip = [System.IO.Compression.ZipFile]::OpenRead($packageFile)
  $entryNames = $zip.Entries | ForEach-Object { $_.FullName.Replace("\", "/") }

  $required = @(
    "[Content_Types].xml",
    "word/vbaProject.bin",
    "customUI/customUI.xml"
  )

  $missing = @()
  foreach ($item in $required) {
    if ($entryNames -notcontains $item) {
      $missing += $item
    }
  }

  if ($missing.Count -gt 0) {
    Write-Error ("Template package is incomplete. Missing: " + ($missing -join ", "))
  }

  $contentTypesEntry = $zip.GetEntry("[Content_Types].xml")
  $relsEntry = $zip.GetEntry("_rels/.rels")
  [xml]$contentTypesXml = (New-Object System.IO.StreamReader($contentTypesEntry.Open(), [System.Text.Encoding]::UTF8)).ReadToEnd()
  [xml]$relsXml = (New-Object System.IO.StreamReader($relsEntry.Open(), [System.Text.Encoding]::UTF8)).ReadToEnd()

  $nsCt = New-Object System.Xml.XmlNamespaceManager($contentTypesXml.NameTable)
  $nsCt.AddNamespace("ct", "http://schemas.openxmlformats.org/package/2006/content-types")
  $override = $contentTypesXml.SelectSingleNode("/ct:Types/ct:Override[@PartName='/customUI/customUI.xml']", $nsCt)
  if ($override -eq $null -or $override.ContentType -ne "application/xml") {
    Write-Error "customUI content type is invalid. Expected application/xml for /customUI/customUI.xml."
  }

  $nsRel = New-Object System.Xml.XmlNamespaceManager($relsXml.NameTable)
  $nsRel.AddNamespace("rel", "http://schemas.openxmlformats.org/package/2006/relationships")
  $relationship = $relsXml.SelectSingleNode("/rel:Relationships/rel:Relationship[@Type='http://schemas.microsoft.com/office/2006/relationships/ui/extensibility']", $nsRel)
  if ($relationship -eq $null -or $relationship.Target -ne "/customUI/customUI.xml") {
    Write-Error "customUI root relationship is invalid. Expected target /customUI/customUI.xml."
  }

  Write-Host "Package validation passed: $packageFile"
  Write-Host "Detected word/vbaProject.bin and customUI/customUI.xml"
  Write-Host "Verified customUI content type=application/xml and relationship target=/customUI/customUI.xml"
}
finally {
  if ($zip -ne $null) {
    $zip.Dispose()
  }
}