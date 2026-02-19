param(
  [string]$SourceDir = "$env:USERPROFILE\Downloads"
)

$ErrorActionPreference = "Stop"
$assetsDir = Join-Path $PSScriptRoot "assets"
if (-not (Test-Path $assetsDir)) {
  New-Item -ItemType Directory -Path $assetsDir | Out-Null
}

$files = Get-ChildItem -Path $SourceDir -File |
  Where-Object {
    @(".png", ".jpg", ".jpeg") -contains $_.Extension.ToLowerInvariant()
  } |
  Sort-Object LastWriteTime -Descending |
  Select-Object -First 5

if (-not $files -or $files.Count -eq 0) {
  Write-Error "No images found in $SourceDir"
}

$index = 1
foreach ($file in $files) {
  $dest = Join-Path $assetsDir ("shot{0}.png" -f $index)
  Copy-Item -Path $file.FullName -Destination $dest -Force
  Write-Host ("Copied {0} -> {1}" -f $file.Name, $dest)
  $index++
}

Write-Host "Done. Imported $($files.Count) screenshot(s)."
