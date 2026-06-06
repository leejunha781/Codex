$ErrorActionPreference = "Stop"
$proposal = Get-ChildItem "D:\이력서" -Directory | Where-Object { $_.Name -like "AVEVA*PLM SME*" } | Select-Object -First 1
$dir = Join-Path $proposal.FullName "Proposal"
$file = Join-Path $dir "Future_Industrial_PLM_Meeting_Deck_EN_Rebuilt_v2.pptx"
$outDir = "C:\Users\namma\.claude\plm_slide_work\hybrid_render"
if (-not (Test-Path $outDir)) { New-Item -ItemType Directory -Path $outDir | Out-Null }

try { $ppt = [Runtime.InteropServices.Marshal]::GetActiveObject("PowerPoint.Application") } catch { $ppt = New-Object -ComObject PowerPoint.Application }
$ppt.Visible = $true
$pres = $null; $opened = $false
foreach ($p in $ppt.Presentations) { if ($p.FullName -eq $file) { $pres = $p; break } }
if (-not $pres) { $pres = $ppt.Presentations.Open($file, $true, $false, $false); $opened = $true }

foreach ($i in 13,15,16,20,21) {
  $out = Join-Path $outDir ("slide$i.png")
  $pres.Slides.Item($i).Export($out, "PNG", 1440, 810)
  Write-Output "RENDERED s$i -> $out  size=$((Get-Item $out).Length)"
}
if ($opened) { $pres.Close() }
Write-Output "DONE"
