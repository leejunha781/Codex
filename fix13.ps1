$ErrorActionPreference = "Stop"
$v2 = "D:\이력서\AVEVA - Marine Principal Technical Support & Consultant – PLM SME, Busan\Proposal\Future_Industrial_PLM_Meeting_Deck_EN_Rebuilt_v2.pptx"
$name = "Future_Industrial_PLM_Meeting_Deck_EN_Rebuilt_v2.pptx"
$outDir = "C:\Users\namma\.claude\plm_slide_work\review_png2"

$ppt = New-Object -ComObject PowerPoint.Application
$pres = $null
foreach($p in $ppt.Presentations){ if($p.Name -eq $name){ $pres=$p; break } }
$openedByUs = $false
if($pres -eq $null){ $pres = $ppt.Presentations.Open($v2, 0, 0, 0); $openedByUs = $true }

$s13 = $pres.Slides.Item(13)
$removed = 0
foreach($sh in @($s13.Shapes)){
    try {
        if($sh.HasTextFrame -eq -1 -and $sh.TextFrame.HasText -eq -1){
            $t = $sh.TextFrame.TextRange.Text
            if($t -like "Sequential delivery*"){ $sh.Delete(); $removed++ }
        }
    } catch {}
}
$pres.Save()
$pres.Slides.Item(13).Export("$outDir\slide-13.png","PNG",1280,720)
if($openedByUs){ $pres.Close() }
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($pres) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($ppt)  | Out-Null
[GC]::Collect(); [GC]::WaitForPendingFinalizers()
Write-Output ("FIX13_DONE removed=" + $removed)
