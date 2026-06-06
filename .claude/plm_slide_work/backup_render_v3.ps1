$ErrorActionPreference = "Stop"
$src = "D:\이력서\AVEVA - Marine Principal Technical Support & Consultant – PLM SME, Busan\Proposal\Future_Industrial_PLM_Meeting_Deck_EN_Rebuilt_v2.pptx"
$bkdir = "D:\이력서\AVEVA - Marine Principal Technical Support & Consultant – PLM SME, Busan\Proposal\_backup_20260606_meeting"
if (-not (Test-Path $bkdir)) { New-Item -ItemType Directory -Path $bkdir | Out-Null }
Copy-Item $src (Join-Path $bkdir "Future_Industrial_PLM_Meeting_Deck_EN_Rebuilt_v2_PREEDIT.pptx") -Force
Write-Output "BACKUP OK"

$rdir = "C:\Users\namma\.claude\plm_slide_work\render_v2"
if (-not (Test-Path $rdir)) { New-Item -ItemType Directory -Path $rdir | Out-Null }

try { $ppt = [Runtime.InteropServices.Marshal]::GetActiveObject("PowerPoint.Application") }
catch { $ppt = New-Object -ComObject PowerPoint.Application }
$pres = $ppt.Presentations.Open($src, $true, $false, $false)
$n = $pres.Slides.Count
for ($i=1; $i -le $n; $i++) {
    $png = Join-Path $rdir ("slide{0:D2}.png" -f $i)
    $pres.Slides.Item($i).Export($png, "PNG", 1280, 720)
}
$pres.Close()
Write-Output "RENDERED $n slides -> $rdir"
