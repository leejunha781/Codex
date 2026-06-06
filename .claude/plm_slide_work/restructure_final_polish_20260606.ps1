$ErrorActionPreference = "Stop"

$deck = "D:\이력서\AVEVA - Marine Principal Technical Support & Consultant – PLM SME, Busan\Proposal\Future_Industrial_PLM_Meeting_Deck_EN_Rebuilt_v2.pptx"
$reviewDir = "C:\Users\namma\.claude\plm_slide_work\restructure_review_20260606"

function Rgb([string]$hex) {
    $hex = $hex.TrimStart('#')
    [int](
        [Convert]::ToInt32($hex.Substring(0,2),16) +
        [Convert]::ToInt32($hex.Substring(2,2),16) * 256 +
        [Convert]::ToInt32($hex.Substring(4,2),16) * 65536
    )
}
function Retry($sb) {
    for ($a=0; $a -lt 8; $a++) {
        try { return (& $sb) } catch { Start-Sleep -Milliseconds 100 }
    }
    return (& $sb)
}

$AMBER = "E8B23A"
$ppt = $null
$pres = $null
try {
    $ppt = New-Object -ComObject PowerPoint.Application
    $pres = $ppt.Presentations.Open($deck,0,0,0)

    function FindSlide($prefix) {
        foreach ($sl in $pres.Slides) {
            foreach ($sh in $sl.Shapes) {
                try {
                    if ($sh.HasTextFrame -eq -1 -and $sh.TextFrame.HasText -eq -1 -and
                        $sh.TextFrame.TextRange.Text.Trim().StartsWith($prefix)) { return $sl }
                } catch {}
            }
        }
        throw ("Slide not found: " + $prefix)
    }

    function SetStarts($sl,$prefix,$newText) {
        foreach ($sh in $sl.Shapes) {
            try {
                if ($sh.HasTextFrame -eq -1 -and $sh.TextFrame.HasText -eq -1 -and
                    $sh.TextFrame.TextRange.Text.Trim().StartsWith($prefix)) {
                    $sh.TextFrame.TextRange.Text = $newText
                    return $sh
                }
            } catch {}
        }
        return $null
    }

    # Current-state table: customer-owned API/process is a proposed strength, not a current full strength.
    $current = FindSlide "Current-state comparison"
    foreach ($sh in $current.Shapes) {
        try {
            if ($sh.HasTextFrame -eq -1 -and $sh.TextFrame.HasText -eq -1 -and
                $sh.TextFrame.TextRange.Text.Trim() -eq "Strong" -and
                $sh.Left -gt 215 -and $sh.Left -lt 245 -and $sh.Top -gt 360 -and $sh.Top -lt 400) {
                $sh.TextFrame.TextRange.Text = "Partial"
                Retry { $sh.TextFrame.TextRange.Font.Color.RGB = [int](Rgb $AMBER) } | Out-Null
            }
        } catch {}
    }

    # Turn the thesis into the explicit meeting topic requested by the user.
    $thesis = FindSlide "Executive thesis"
    SetStarts $thesis "Executive thesis" "Meeting thesis - win above tool silos" | Out-Null
    SetStarts $thesis "The future PLM opportunity" "Own the change-to-assurance-to-operations loop, not every specialist calculation" | Out-Null
    SetStarts $thesis "North Star" "North Star`rRequirement -> Engineering -> Assure -> BOM -> Build -> Commissioning -> Handover -> Operations -> Maintenance -> Change Feedback" | Out-Null
    SetStarts $thesis "Strategic differentiation" "Strategic differentiation`rA vendor-neutral Digital Thread Hub + Continuous Naval Assurance Control Plane with configuration governance, executable APIs, approved-solver evidence and measurable quality gates." | Out-Null

    # Clearly distinguish what is executable today from the proposed pilot extension.
    $reference = FindSlide "Executable reference package"
    $gov = SetStarts $reference "Governance value" "Governance value`rThe current executable package proves API-first contracts, security and gate patterns. Naval Assurance adapters, evidence lineage and operations calibration are the proposed pilot extension."
    if ($gov -ne $null) { $gov.TextFrame.TextRange.Font.Size = [single]9.5 }
    SetStarts $reference "Meeting objective" "Meeting objective: agree the executable baseline, then approve the bounded Continuous Naval Assurance pilot extension." | Out-Null

    # Restore the visual hierarchy inside the architecture layer boxes.
    $arch = FindSlide "Target architecture concept"
    foreach ($sh in $arch.Shapes) {
        try {
            if ($sh.HasTextFrame -eq -1 -and $sh.TextFrame.HasText -eq -1) {
                $t = $sh.TextFrame.TextRange.Text.Trim()
                if ($t.Contains("Layer") -and $t.Contains("`r")) {
                    $sh.TextFrame.TextRange.Paragraphs(1,1).Font.Bold = [int]-1
                    $sh.TextFrame.TextRange.Paragraphs(1,1).Font.Size = [single]11
                    $sh.TextFrame.TextRange.Paragraphs(2,1).Font.Bold = [int]0
                    $sh.TextFrame.TextRange.Paragraphs(2,1).Font.Size = [single]10.5
                }
            }
        } catch {}
    }

    $pres.Save()
    foreach ($idx in @(3,7,11,12)) {
        $pres.Slides.Item($idx).Export((Join-Path $reviewDir ("slide-{0:D2}.png" -f $idx)),"PNG",1280,720)
    }
    Write-Output "FINAL_POLISH_DONE"
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
