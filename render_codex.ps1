$ErrorActionPreference = "Stop"
$src = "D:\이력서\AVEVA - Marine Principal Technical Support & Consultant – PLM SME, Busan\Proposal\Future_Industrial_PLM_Meeting_Deck_EN_Rebuilt_v2.pptx"
$rdir = "C:\Users\namma\.claude\plm_slide_work\render_v2"
$ppt = New-Object -ComObject PowerPoint.Application
$pres = $ppt.Presentations.Open($src, $true, $false, $false)
foreach($i in 2,18,31){ $pres.Slides.Item($i).Export((Join-Path $rdir ("codex{0:D2}.png" -f $i)), "PNG", 1600, 900) }
$pres.Close()
if($ppt.Presentations.Count -eq 0){$ppt.Quit()}
"rendered 2,18,31"
