$ErrorActionPreference = "Stop"
$src = "D:\이력서\AVEVA - Marine Principal Technical Support & Consultant – PLM SME, Busan\Proposal\Future_Industrial_PLM_Meeting_Deck_EN_Rebuilt_v2.pptx"
$rdir = "C:\Users\namma\.claude\plm_slide_work\render_v2"
try { $ppt = [Runtime.InteropServices.Marshal]::GetActiveObject("PowerPoint.Application") }
catch { $ppt = New-Object -ComObject PowerPoint.Application }
$pres = $ppt.Presentations.Open($src, $true, $false, $false)
$pres.Slides.Item(31).Export((Join-Path $rdir "slide31.png"), "PNG", 1600, 900)
$pres.Slides.Item(32).Export((Join-Path $rdir "slide32.png"), "PNG", 1600, 900)
$pres.Close()
Write-Output "RENDERED 31,32"
