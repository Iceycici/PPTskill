# render_pages.ps1 - render selected slides of a pptx to PNG (COM, 1600x900)
# Usage:
#   powershell -NoProfile -ExecutionPolicy Bypass -File render_pages.ps1 -Pptx "C:\path\deck.pptx" -OutDir "C:\path\out" -Pages 2,31,58
# Notes:
#   - Output name: s<page>.png (no zero padding). Use render_all.ps1 for full export.
#   - ASCII-only on purpose: PowerShell 5.1 reads BOM-less scripts in the local
#     ANSI codepage; non-ASCII comments/paths corrupt parsing (see skill references/04).
#   - If your file name contains non-ASCII characters, resolve it with Get-ChildItem
#     on the caller side and pass the resolved FullName via -Pptx.
# Source: generalized from the 2026-09-17 zcode session workflow (deck/render_v4b.ps1).
param(
  [Parameter(Mandatory = $true)][string]$Pptx,
  [Parameter(Mandatory = $true)][string]$OutDir,
  [Parameter(Mandatory = $true)][int[]]$Pages
)
$ErrorActionPreference = "Stop"

if (-not (Test-Path $Pptx)) { throw ("File not found: " + $Pptx) }
$src = (Resolve-Path $Pptx).Path
New-Item -ItemType Directory -Force -Path $OutDir | Out-Null

$pp = New-Object -ComObject PowerPoint.Application
try {
  $pres = $pp.Presentations.Open($src, $true, $false, $false)   # ReadOnly, no window
  foreach ($i in $Pages) {
    $pres.Slides.Item($i).Export((Join-Path $OutDir ("s" + $i + ".png")), "PNG", 1600, 900)
  }
  $pres.Close()
  Write-Output ("RENDERED " + $Pages.Count + " pages -> " + $OutDir)
}
finally {
  $pp.Quit()
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($pp) | Out-Null
}
