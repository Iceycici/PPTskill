# render_all.ps1 - render every slide of a pptx to PNG (COM, 1600x900)
# Usage:
#   powershell -NoProfile -ExecutionPolicy Bypass -File render_all.ps1 -Pptx "C:\path\deck.pptx" -OutDir "C:\path\render"
# Notes:
#   - Output: s01.png / s02.png ... (zero-padded), 1600x900. Requires PowerPoint or WPS (COM).
#   - ASCII-only on purpose: PowerShell 5.1 reads BOM-less scripts in the local ANSI
#     codepage; non-ASCII comments/paths corrupt parsing (see skill references/04).
#   - If the file name contains non-ASCII characters, resolve it on the caller side:
#       $src = (Get-ChildItem "C:\dir" -Filter "AI*.pptx" | Select-Object -First 1).FullName
#     (the wildcard trick also excludes Office lock files ~$*).
#   - If the file is open in Office (a ~$ lock file exists), close it first.
# Source: generalized from the 2026-09-17 zcode session workflow (deck/export_png.ps1).
param(
  [Parameter(Mandatory = $true)][string]$Pptx,
  [Parameter(Mandatory = $true)][string]$OutDir
)
$ErrorActionPreference = "Stop"

if (-not (Test-Path $Pptx)) { throw ("File not found: " + $Pptx) }
$src = (Resolve-Path $Pptx).Path
New-Item -ItemType Directory -Force -Path $OutDir | Out-Null
Get-ChildItem $OutDir -Filter *.png -ErrorAction SilentlyContinue | Remove-Item -Force

$pp = New-Object -ComObject PowerPoint.Application
try {
  $pres = $pp.Presentations.Open($src, $true, $false, $false)   # ReadOnly, no window
  $n = $pres.Slides.Count
  for ($i = 1; $i -le $n; $i++) {
    $name = "s{0:D2}.png" -f $i
    $pres.Slides.Item($i).Export((Join-Path $OutDir $name), "PNG", 1600, 900)
  }
  $pres.Close()
  Write-Output ("RENDERED " + $n + " pages -> " + $OutDir)
}
finally {
  $pp.Quit()
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($pp) | Out-Null
}
