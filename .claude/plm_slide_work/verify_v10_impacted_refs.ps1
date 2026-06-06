$ErrorActionPreference = 'Stop'

$deckPath = 'D:\이력서\AVEVA - Marine Principal Technical Support & Consultant – PLM SME, Busan\Proposal\Future_Industrial_PLM_Meeting_Deck_EN_V10.pptx'
$reportPath = 'C:\Users\namma\.claude\plm_slide_work\v10_impacted_refs_verify.txt'

function Get-PowerPointApplication {
    try {
        return [System.Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application')
    } catch {
        try {
            return (New-Object -ComObject PowerPoint.Application)
        } catch {
            $exe = 'C:\Program Files\Microsoft Office\root\Office16\POWERPNT.EXE'
            Start-Process -FilePath $exe -WindowStyle Hidden
            Start-Sleep -Seconds 8
            return [System.Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application')
        }
    }
}

function Get-ShapeTexts {
    param($Shapes)
    $items = New-Object System.Collections.Generic.List[string]
    foreach ($shape in $Shapes) {
        try {
            if ($shape.Type -eq 6) {
                $childTexts = Get-ShapeTexts -Shapes $shape.GroupItems
                foreach ($t in $childTexts) { [void]$items.Add($t) }
            } elseif ($shape.HasTextFrame -and $shape.TextFrame.HasText) {
                $txt = $shape.TextFrame.TextRange.Text
                $txt = ($txt -replace "`r", " " -replace "`v", " " -replace "\s+", " ").Trim()
                if ($txt.Length -gt 0) { [void]$items.Add($txt) }
            }
        } catch {}
    }
    return $items
}

$ppt = $null
$pres = $null
try {
    $ppt = Get-PowerPointApplication
    $pres = $ppt.Presentations.Open($deckPath, $true, $false, $false)
    $lines = New-Object System.Collections.Generic.List[string]
    [void]$lines.Add('V10 impacted reference verification')
    [void]$lines.Add("Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss zzz')")
    [void]$lines.Add("Deck: $deckPath")
    [void]$lines.Add("Slides: $($pres.Slides.Count)")
    [void]$lines.Add('')

    $oldMatches = New-Object System.Collections.Generic.List[string]
    for ($i = 1; $i -le $pres.Slides.Count; $i++) {
        $texts = Get-ShapeTexts -Shapes $pres.Slides.Item($i).Shapes
        foreach ($t in $texts) {
            if ($t -match 'p\.28|p\.3-5|Meeting Procedure') {
                [void]$oldMatches.Add("SLIDE ${i}: $t")
            }
        }
    }
    [void]$lines.Add('Targeted matches for p.28 / p.3-5 / Meeting Procedure:')
    if ($oldMatches.Count -eq 0) {
        [void]$lines.Add('NONE')
    } else {
        foreach ($m in $oldMatches) { [void]$lines.Add($m) }
    }
    [void]$lines.Add('')

    foreach ($i in @(1,2,18,38)) {
        [void]$lines.Add("Slide $i checked text:")
        foreach ($t in (Get-ShapeTexts -Shapes $pres.Slides.Item($i).Shapes)) {
            if ($t -match 'Rev\.|p\.|Pilot value:|Future Industrial PLM /|Meeting procedure guide|Glossary appendix|WBS') {
                [void]$lines.Add("  $t")
            }
        }
        [void]$lines.Add('')
    }

    [System.IO.File]::WriteAllLines($reportPath, $lines.ToArray(), [System.Text.Encoding]::UTF8)
    Write-Host ([string]::Join("`n", $lines.ToArray()))
} finally {
    if ($pres -ne $null) {
        try { $pres.Close() } catch {}
    }
    if ($ppt -ne $null) {
        try {
            if ($ppt.Presentations.Count -eq 0) { $ppt.Quit() }
        } catch {}
    }
    if ($pres -ne $null) {
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($pres) | Out-Null 2>$null
    }
    if ($ppt -ne $null) {
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ppt) | Out-Null 2>$null
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
