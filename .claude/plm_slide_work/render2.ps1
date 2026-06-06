$ErrorActionPreference = "Stop"
$proposal = Get-ChildItem "D:\이력서" -Directory | Where-Object { $_.Name -like "AVEVA*PLM SME*" } | Select-Object -First 1
$file = Join-Path $proposal.FullName "Proposal\Future_Industrial_PLM_Meeting_Deck_EN_Rebuilt_v2.pptx"
$outDir = "C:\Users\namma\.claude\plm_slide_work\hybrid_render"
try { $ppt = [Runtime.InteropServices.Marshal]::GetActiveObject("PowerPoint.Application") } catch { $ppt = New-Object -ComObject PowerPoint.Application }
$ppt.Visible = $true
$pres = $null
foreach ($p in $ppt.Presentations) { if ($p.FullName -eq $file) { $pres = $p; break } }
$opened=$false
if (-not $pres) { $pres = $ppt.Presentations.Open($file, $true, $false, $false); $opened=$true }
foreach ($i in 13,18) { $out=Join-Path $outDir ("slide$i.png"); $pres.Slides.Item($i).Export($out,"PNG",1440,810); Write-Output "RENDERED s$i" }
if ($opened) { $pres.Close() }
Write-Output "DONE"
