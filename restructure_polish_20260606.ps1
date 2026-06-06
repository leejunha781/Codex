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

$MUTE = "93A8C0"
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

    function SetContaining($sl,$needle,$newText) {
        foreach ($sh in $sl.Shapes) {
            try {
                if ($sh.HasTextFrame -eq -1 -and $sh.TextFrame.HasText -eq -1 -and
                    $sh.TextFrame.TextRange.Text.Contains($needle)) {
                    $sh.TextFrame.TextRange.Text = $newText
                    return $sh
                }
            } catch {}
        }
        return $null
    }

    $gap = FindSlide "Strategic gap to fill"
    SetContaining $gap "The missing product" "The missing product`rA configuration-governed Digital Thread + Continuous Naval Assurance Control Plane that connects engineering changes to approved-solver evidence, approvals and operations feedback." | Out-Null
    SetContaining $gap "Why teams should agree" "Why teams should agree`rIt complements NAPA/class solvers and existing PLM/CAD/MES tools instead of replacing them first. This lowers adoption and class-acceptance risk while creating visible value early." | Out-Null

    $arch = FindSlide "Target architecture concept"
    SetContaining $arch "Experience Layer" "Experience Layer`rReact dashboard, 3D context, assurance margin view, gate report and collaboration views" | Out-Null
    SetContaining $arch "API Contract Layer" "API Contract Layer`rFastAPI, OpenAPI, JWT/OIDC, role-based promote authority and solver-adapter contracts" | Out-Null
    SetContaining $arch "Assurance + Digital Thread Layer" "Assurance + Digital Thread Layer`rPostgreSQL + assurance graph + rule engine + evaluate/promote flow + evidence lineage" | Out-Null
    SetContaining $arch "Integration Layer" "Integration Layer`rAdapters for E3D, NAPA/class/CFD/FEM, PLM, ERP, MES, AIM and PI / CONNECT" | Out-Null

    $phase = FindSlide "Phase 2 + Continuous Naval Assurance"
    SetContaining $phase "Design rule" "Design rule`rConfiguration is the control layer that makes handover, operations and naval-assurance impact analysis reliable without replacing certified solvers." | Out-Null
    SetContaining $phase "Phase 2 output" "Phase 2 output`rA governed configuration + assurance graph where each change links to affected objects, loading cases, solver runs, evidence and approvals." | Out-Null

    $decision = FindSlide "Recommended decision"
    foreach ($sh in $decision.Shapes) {
        try {
            if ($sh.HasTextFrame -eq -1 -and $sh.TextFrame.HasText -eq -1) {
                $t = $sh.TextFrame.TextRange.Text.Trim()
                if ($t.StartsWith("Proceed with Phase 2")) {
                    $sh.TextFrame.TextRange.Font.Size = [single]14.5
                    $sh.TextFrame.MarginTop = [single]6
                    $sh.TextFrame.MarginBottom = [single]6
                }
                if ($t.StartsWith("Why now")) {
                    $sh.TextFrame.TextRange.Font.Size = [single]13.5
                    $sh.TextFrame.MarginTop = [single]5
                    $sh.TextFrame.MarginBottom = [single]5
                }
            }
        } catch {}
    }

    # Standardize every visible page number to the original deck location.
    foreach ($sl in $pres.Slides) {
        $found = $false
        foreach ($sh in $sl.Shapes) {
            try {
                if ($sh.HasTextFrame -eq -1 -and $sh.TextFrame.HasText -eq -1 -and
                    $sh.TextFrame.TextRange.Text.Trim() -eq [string]$sl.SlideIndex -and
                    $sh.Left -gt 880 -and $sh.Top -lt 60) {
                    $sh.Left = [single]920
                    $sh.Top = [single]9
                    $sh.Width = [single]18
                    $sh.Height = [single]12
                    $sh.TextFrame.MarginLeft = [single]0
                    $sh.TextFrame.MarginRight = [single]0
                    $sh.TextFrame.MarginTop = [single]0
                    $sh.TextFrame.MarginBottom = [single]0
                    $sh.TextFrame.TextRange.Font.Size = [single]6.5
                    $sh.TextFrame.TextRange.Font.Bold = [int]0
                    $sh.TextFrame.TextRange.Font.Name = "Aptos"
                    Retry { $sh.TextFrame.TextRange.Font.Color.RGB = [int](Rgb $MUTE) } | Out-Null
                    $sh.TextFrame.TextRange.ParagraphFormat.Alignment = [int]3
                    $found = $true
                }
            } catch {}
        }
    }

    $pres.Save()
    foreach ($idx in @(8,11,14,22,23,24,25,26,27,28)) {
        $pres.Slides.Item($idx).Export((Join-Path $reviewDir ("slide-{0:D2}.png" -f $idx)),"PNG",1280,720)
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
