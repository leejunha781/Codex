$ErrorActionPreference = "Stop"

$deck = "D:\이력서\AVEVA - Marine Principal Technical Support & Consultant – PLM SME, Busan\Proposal\Future_Industrial_PLM_Meeting_Deck_EN_Rebuilt_v2.pptx"
$reviewDir = "C:\Users\namma\.claude\plm_slide_work\restructure_review_20260606"
$ppt = $null
$pres = $null
try {
    $ppt = New-Object -ComObject PowerPoint.Application
    $pres = $ppt.Presentations.Open($deck,0,0,0)
    $fixes = @(
        @{s=7;  n="Text 6";     h=86},
        @{s=13; n="TextBox 56"; h=105},
        @{s=14; n="Text 1";     h=40},
        @{s=17; n="TextBox 60"; h=100},
        @{s=18; n="TextBox 54"; h=82},
        @{s=18; n="TextBox 61"; h=82},
        @{s=20; n="TextBox 69"; h=120}
    )
    foreach ($f in $fixes) {
        $sh = $pres.Slides.Item([int]$f.s).Shapes.Item([string]$f.n)
        $sh.Height = [single]$f.h
    }
    $pres.Save()
    foreach ($idx in @(7,13,14,17,18,20)) {
        $pres.Slides.Item($idx).Export((Join-Path $reviewDir ("slide-{0:D2}.png" -f $idx)),"PNG",1280,720)
    }
    Write-Output "TEXT_BOUNDS_FIXED"
}
finally {
    if ($pres -ne $null) {
        try { $pres.Close() } catch {}
        try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($pres) | Out-Null } catch {}
    }
    if ($ppt -ne $null) {
        try { if ($ppt.Presentations.Count -eq 0) { $ppt.Quit() } } catch {}
        try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ppt) | Out-Null } catch {}
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
