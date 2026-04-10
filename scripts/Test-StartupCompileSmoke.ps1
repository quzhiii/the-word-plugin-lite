$ErrorActionPreference = "Stop"

$word = $null

try {
  $word = New-Object -ComObject Word.Application
  $word.Visible = $false
  $word.DisplayAlerts = 0

  $result = $word.Run("THU_Formatter_Addin.AddinSelfCheck")
  if ($result -notlike "BRIDGE_VERSION=*") {
    throw "Unexpected self-check output: $result"
  }

  Write-Host $result
}
finally {
  if ($null -ne $word) {
    $word.Quit()
  }
}
