param(
  [int]$Duration = 30,
  [int]$SetupDelay = 8,
  [int]$Fps = 30,
  [string]$Output = "quickcanvas-ad.mp4",
  [switch]$CaptureDesktop,
  [switch]$CaptureWindow
)

$ErrorActionPreference = "Stop"
$indexPath = (Resolve-Path (Join-Path $PSScriptRoot "index.html")).Path
$indexUrl = "file:///" + ($indexPath -replace "\\", "/")
$captureInput = "desktop"
if ($CaptureWindow) {
  $captureInput = "title=QuickCanvas Ad Kit"
}

Write-Host "Opening ad page..."
$edgeArgs = @(
  "--new-window",
  "--start-fullscreen",
  """$indexUrl""",
  "--disable-gpu",
  "--disable-gpu-compositing",
  "--disable-features=UseSkiaRenderer"
) -join " "
Start-Process "msedge.exe" $edgeArgs
Start-Sleep -Seconds $SetupDelay

Write-Host "Recording $Duration second(s) to $Output..."
$ffmpegArgs = @(
  "-y",
  "-f", "gdigrab",
  "-framerate", "$Fps",
  "-draw_mouse", "0",
  "-i", $captureInput,
  "-t", "$Duration",
  "-vf", "crop='if(gte(iw/ih,9/16),ih*9/16,iw)':'if(gte(iw/ih,9/16),ih,iw*16/9)',scale=1080:1920",
  "-c:v", "libx264",
  "-preset", "medium",
  "-crf", "20",
  "-pix_fmt", "yuv420p",
  "-movflags", "+faststart",
  "$Output"
)

& ffmpeg @ffmpegArgs

if ($LASTEXITCODE -eq 0 -and (Test-Path $Output) -and (Get-Item $Output).Length -gt 0) {
  Write-Host "Done. MP4 created: $Output"
} else {
  Write-Error "Recording failed. Ensure the Edge app window title is 'QuickCanvas Ad Kit'."
}
