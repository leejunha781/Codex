$ErrorActionPreference = "Stop"
$dir  = "D:\이력서\AVEVA - Marine Principal Technical Support & Consultant – PLM SME, Busan\Proposal"
$orig = Join-Path $dir "Future_Industrial_PLM_Meeting_Deck_EN_Rebuilt.pptx"
$v2   = Join-Path $dir "Future_Industrial_PLM_Meeting_Deck_EN_Rebuilt_v2.pptx"
$png  = "C:\Users\namma\.claude\plm_slide_work\orig_slide19.png"
$origName = "Future_Industrial_PLM_Meeting_Deck_EN_Rebuilt.pptx"

$ppt = New-Object -ComObject PowerPoint.Application

# Find the original if it's already open in the (user's) running instance.
$pres = $null
foreach($p in $ppt.Presentations){ if($p.Name -eq $origName){ $pres = $p; break } }
$openedByUs = $false
if($pres -eq $null){
    $pres = $ppt.Presentations.Open($orig, 0, 0, 0)  # windowless
    $openedByUs = $true
    Write-Output "Original was CLOSED -> opened windowless"
} else {
    Write-Output "Original is OPEN in the running instance -> inserting into it"
}

$before = [int]$pres.Slides.Count
# Insert ONLY slide 19 from v2, appended after the last slide. v2 is a copy of the
# original (same theme), so the inserted slide is visually identical to the QA render.
$pres.Slides.InsertFromFile($v2, $before, 19, 19) | Out-Null
$after = [int]$pres.Slides.Count

$pres.Save()
$pres.Slides.Item($after).Export($png, "PNG", 1280, 720)

if($openedByUs){ $pres.Close() }   # never $ppt.Quit() - may be the user's shared instance

[System.Runtime.InteropServices.Marshal]::ReleaseComObject($pres) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($ppt)  | Out-Null
[GC]::Collect(); [GC]::WaitForPendingFinalizers()
Write-Output "MERGED before=$before after=$after into ORIGINAL"
