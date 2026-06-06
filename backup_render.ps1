$ErrorActionPreference = "Stop"
$dir = "D:\이력서\AVEVA - Marine Principal Technical Support & Consultant – PLM SME, Busan\Proposal"
$v2 = Join-Path $dir "Future_Industrial_PLM_Meeting_Deck_EN_Rebuilt_v2.pptx"
$name = "Future_Industrial_PLM_Meeting_Deck_EN_Rebuilt_v2.pptx"
$outDir = "C:\Users\namma\.claude\plm_slide_work\review_full"
New-Item -ItemType Directory -Force -Path $outDir | Out-Null
Get-ChildItem -LiteralPath $outDir -Filter *.png | Remove-Item -Force -ErrorAction SilentlyContinue

# ---- (3) backup v2 BEFORE any write ----
$ts = Get-Date -Format "yyyyMMdd_HHmmss"
$bk = Join-Path $dir ("Future_Industrial_PLM_Meeting_Deck_EN_Rebuilt_v2_BACKUP_$ts.pptx")
Copy-Item -LiteralPath $v2 -Destination $bk -Force
Write-Output ("V2_BACKUP=" + $bk)

# ---- render all slides (read-only) ----
$ppt = New-Object -ComObject PowerPoint.Application
$pres=$null
foreach($p in @($ppt.Presentations)){ if($p.Name -eq $name){ $pres=$p; break } }
$byUs=$false
if($pres -eq $null){ $pres=$ppt.Presentations.Open($v2,-1,0,0); $byUs=$true }
$n=$pres.Slides.Count
Write-Output ("SLIDES=" + $n)
for($i=1;$i -le $n;$i++){ $num="{0:D2}" -f $i; $pres.Slides.Item($i).Export("$outDir\slide-$num.png","PNG",1280,720) }
if($byUs){ $pres.Close() }
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($pres) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($ppt) | Out-Null
[GC]::Collect()
Write-Output "RENDER_DONE"
