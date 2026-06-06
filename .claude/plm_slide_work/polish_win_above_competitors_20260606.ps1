$ErrorActionPreference = "Stop"

$deck = "D:\이력서\AVEVA - Marine Principal Technical Support & Consultant – PLM SME, Busan\Proposal\Future_Industrial_PLM_Meeting_Deck_EN_Rebuilt_v2.pptx"
$reviewDir = "C:\Users\namma\.claude\plm_slide_work\win_above_review_20260606"
$ppt = $null
$pres = $null

try {
    $ppt = New-Object -ComObject PowerPoint.Application
    $pres = $ppt.Presentations.Open($deck,0,0,0)

    function GetSlide($prefix) {
        foreach ($sl in $pres.Slides) {
            foreach ($sh in $sl.Shapes) {
                try {
                    if ($sh.HasTextFrame -eq -1 -and $sh.TextFrame.HasText -eq -1 -and
                        $sh.TextFrame.TextRange.Text.Trim().StartsWith($prefix)) {
                        return $sl
                    }
                } catch {}
            }
        }
        throw ("Slide not found: " + $prefix)
    }

    function ReplaceContaining($sl,$pattern,$newText) {
        foreach ($sh in $sl.Shapes) {
            try {
                if ($sh.HasTextFrame -eq -1 -and $sh.TextFrame.HasText -eq -1 -and
                    $sh.TextFrame.TextRange.Text.Contains($pattern)) {
                    $sh.TextFrame.TextRange.Text = [string]$newText
                    return $sh
                }
            } catch {}
        }
        throw ("Text not found: " + $pattern)
    }

    $sThesis = $pres.Slides.Item(7)
    $thesis = $sThesis.Shapes.Item("Text 6")
    $thesis.TextFrame.TextRange.Text = "Strategic differentiation`rOpen AVEVA control plane = Siemens-grade configuration + Dassault-grade collaboration + NAPA-grade assurance, connected by customer-owned APIs, evidence and operations feedback."
    $thesis.TextFrame.TextRange.Font.Name = "Aptos"
    $thesis.TextFrame.TextRange.Font.Size = [single]10
    $thesis.TextFrame.TextRange.Font.Bold = [int]-1

    $sArchitecture = $pres.Slides.Item(13)
    $subtitle = $sArchitecture.Shapes.Item("Text 1")
    $subtitle.TextFrame.TextRange.Text = "The three win-above patterns converge into one open, customer-owned AVEVA control plane"
    $subtitle.TextFrame.TextRange.Font.Name = "Aptos"
    $subtitle.TextFrame.TextRange.Font.Size = [single]11
    $subtitle.TextFrame.TextRange.Font.Bold = [int]0

    $pres.Save()
    foreach ($sl in $pres.Slides) {
        $sl.Export((Join-Path $reviewDir ("slide-{0:D2}.png" -f $sl.SlideIndex)),"PNG",1280,720)
    }
    Write-Output "POLISH_DONE"
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
