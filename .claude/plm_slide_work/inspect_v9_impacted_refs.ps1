$ErrorActionPreference = 'Stop'

$deckPath = 'D:\이력서\AVEVA - Marine Principal Technical Support & Consultant – PLM SME, Busan\Proposal\Future_Industrial_PLM_Meeting_Deck_EN_V9.pptx'
$workDir = 'C:\Users\namma\.claude\plm_slide_work'
$outPath = Join-Path $workDir 'v9_impacted_page_reference_audit.txt'
$fullTextPath = Join-Path $workDir 'v9_full_slide_text_dump.txt'

function Get-ShapeTexts {
    param($Shapes)
    $items = New-Object System.Collections.Generic.List[string]
    foreach ($shape in $Shapes) {
        try {
            if ($shape.Type -eq 6) {
                foreach ($child in $shape.GroupItems) {
                    $childTexts = Get-ShapeTexts -Shapes @($child)
                    foreach ($t in $childTexts) { [void]$items.Add($t) }
                }
            } else {
                if ($shape.HasTextFrame -and $shape.TextFrame.HasText) {
                    $txt = $shape.TextFrame.TextRange.Text
                    $txt = ($txt -replace "`r", "`n" -replace "`v", "`n" -replace "\s+", " ").Trim()
                    if ($txt.Length -gt 0) { [void]$items.Add($txt) }
                }
                if ($shape.HasTable) {
                    $table = $shape.Table
                    for ($r = 1; $r -le $table.Rows.Count; $r++) {
                        for ($c = 1; $c -le $table.Columns.Count; $c++) {
                            try {
                                $cellText = $table.Cell($r, $c).Shape.TextFrame.TextRange.Text
                                $cellText = ($cellText -replace "`r", " " -replace "`v", " " -replace "\s+", " ").Trim()
                                if ($cellText.Length -gt 0) { [void]$items.Add($cellText) }
                            } catch {}
                        }
                    }
                }
            }
        } catch {
            [void]$items.Add("[shape read error: $($_.Exception.Message)]")
        }
    }
    return $items
}

function Get-NotesTexts {
    param($Slide)
    $items = New-Object System.Collections.Generic.List[string]
    try {
        foreach ($shape in $Slide.NotesPage.Shapes) {
            try {
                if ($shape.HasTextFrame -and $shape.TextFrame.HasText) {
                    $txt = $shape.TextFrame.TextRange.Text
                    $txt = ($txt -replace "`r", "`n" -replace "`v", "`n" -replace "\s+", " ").Trim()
                    if ($txt.Length -gt 0) { [void]$items.Add($txt) }
                }
            } catch {}
        }
    } catch {}
    return $items
}

$ppt = $null
$pres = $null
try {
    $ppt = New-Object -ComObject PowerPoint.Application
    $pres = $ppt.Presentations.Open($deckPath, $true, $false, $false)

    $audit = New-Object System.Collections.Generic.List[string]
    $dump = New-Object System.Collections.Generic.List[string]
    [void]$audit.Add("V9 impacted page/reference audit")
    [void]$audit.Add("Deck: $deckPath")
    [void]$audit.Add("Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss zzz')")
    [void]$audit.Add("Slides: $($pres.Slides.Count)")
    [void]$audit.Add("")

    $refPattern = '(?i)(\bp\.?\s*\d+|\bpage\s+\d+|\bpages\s+\d+|\bslide\s+\d+|\bslides\s+\d+|\bp\.\s*\d+\s*[-–]\s*\d+|\bp\.\s*\d+\s*[&,+]\s*\d+|\bp\.\s*\d+\s*[-–]\s*\d+|\bp\.\s*\d+\s*[-–]\s*\d+)'

    foreach ($slide in $pres.Slides) {
        $idx = [int]$slide.SlideIndex
        $texts = Get-ShapeTexts -Shapes $slide.Shapes
        $notes = Get-NotesTexts -Slide $slide

        $title = ''
        foreach ($t in $texts) {
            if ($t.Length -gt 0) { $title = $t; break }
        }
        if ($title.Length -gt 90) { $title = $title.Substring(0, 90) + '...' }
        [void]$dump.Add("===== SLIDE $idx | $title =====")
        foreach ($t in $texts) { [void]$dump.Add("TEXT: $t") }
        foreach ($n in $notes) { [void]$dump.Add("NOTE: $n") }
        [void]$dump.Add("")

        foreach ($t in $texts) {
            if ($t -match $refPattern -or $t -match '(?i)glossary|agenda|procedure|appendix|decision|close|review|p\.\d|p\d') {
                [void]$audit.Add("SLIDE $idx TEXT: $t")
            }
        }
        foreach ($n in $notes) {
            if ($n -match $refPattern -or $n -match '(?i)glossary|agenda|procedure|appendix|decision|close|review|p\.\d|p\d') {
                [void]$audit.Add("SLIDE $idx NOTE: $n")
            }
        }

        $numericTexts = @()
        foreach ($t in $texts) {
            $clean = ($t -replace '\s+', '').Trim()
            if ($clean -match '^\d{1,2}$') { $numericTexts += $clean }
        }
        foreach ($num in $numericTexts) {
            if ([int]$num -ne $idx) {
                [void]$audit.Add("SLIDE $idx NUMERIC-MISMATCH-CANDIDATE: '$num'")
            }
        }
    }

    [System.IO.File]::WriteAllLines($outPath, $audit.ToArray(), [System.Text.Encoding]::UTF8)
    [System.IO.File]::WriteAllLines($fullTextPath, $dump.ToArray(), [System.Text.Encoding]::UTF8)

    Write-Host "Audit: $outPath"
    Write-Host "Full text: $fullTextPath"
} finally {
    if ($pres -ne $null) {
        try { $pres.Close() } catch {}
    }
    if ($ppt -ne $null) {
        try { $ppt.Quit() } catch {}
    }
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($pres) | Out-Null 2>$null
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ppt) | Out-Null 2>$null
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
