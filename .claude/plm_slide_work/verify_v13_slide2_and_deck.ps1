$ErrorActionPreference = 'Stop'

$deck = 'D:\이력서\AVEVA - Marine Principal Technical Support & Consultant – PLM SME, Busan\Proposal\Future_Industrial_PLM_Meeting_Deck_EN_V13.pptx'
$out = 'C:\Users\namma\.claude\plm_slide_work\v13_slide2_and_deck_verify.txt'

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

function Get-ShapeText {
    param($Shape)
    try {
        if ($Shape.HasTextFrame -and $Shape.TextFrame.HasText) {
            return (($Shape.TextFrame.TextRange.Text -replace "`r", " " -replace "`v", " " -replace "\s+", " ").Trim())
        }
    } catch {}
    return ''
}

$ppt = $null
$pres = $null
try {
    $ppt = Get-PowerPointApplication
    $pres = $ppt.Presentations.Open($deck, $true, $false, $false)
    $lines = New-Object System.Collections.Generic.List[string]
    [void]$lines.Add('V13 slide 2 and deck verification')
    [void]$lines.Add("Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss zzz')")
    [void]$lines.Add("Deck: $deck")
    [void]$lines.Add("Slides: $($pres.Slides.Count)")

    $slide = $pres.Slides.Item(2)
    $oldNames = @('GlossaryPointer', 'PresenterGuidePointer')
    $newNames = @(
        'ReferenceToolbarPanelV13',
        'ProcedureGuideCardV13',
        'ProcedureGuideAccentBarV13',
        'ProcedureGuideLabelV13',
        'ProcedureGuideRangeV13',
        'GlossaryGuideCardV13',
        'GlossaryGuideAccentBarV13',
        'GlossaryGuideLabelV13',
        'GlossaryGuideRangeV13'
    )
    [void]$lines.Add('')
    [void]$lines.Add('Slide 2 old pointer shapes:')
    foreach ($name in $oldNames) {
        $found = $false
        foreach ($shape in $slide.Shapes) {
            if ($shape.Name -eq $name) { $found = $true }
        }
        [void]$lines.Add("  $name = $found")
    }
    [void]$lines.Add('')
    [void]$lines.Add('Slide 2 new reference-card shapes:')
    foreach ($name in $newNames) {
        $found = $false
        foreach ($shape in $slide.Shapes) {
            if ($shape.Name -eq $name) { $found = $true }
        }
        [void]$lines.Add("  $name = $found")
    }
    [void]$lines.Add('')
    [void]$lines.Add('Slide 2 reference texts:')
    foreach ($shape in $slide.Shapes) {
        $txt = Get-ShapeText -Shape $shape
        if ($txt -match 'MEETING PROCEDURE|Guide\s+p\.3|ABBREVIATIONS|Glossary\s+p\.39') {
            [void]$lines.Add(("  {0}: {1}" -f $shape.Name, $txt))
        }
    }
    [void]$lines.Add('')
    [void]$lines.Add('Page number text check:')
    foreach ($i in @(1,2,3,4,5,38,39,40,41)) {
        $numeric = New-Object System.Collections.Generic.List[string]
        foreach ($shape in $pres.Slides.Item($i).Shapes) {
            $txt = Get-ShapeText -Shape $shape
            if (($txt -replace '\s+', '') -match '^\d{1,2}$') {
                [void]$numeric.Add($txt)
            }
        }
        [void]$lines.Add(("  Slide {0}: {1}" -f $i, ([string]::Join(', ', $numeric.ToArray()))))
    }

    [System.IO.File]::WriteAllLines($out, $lines.ToArray(), [System.Text.Encoding]::UTF8)
    Write-Host ([string]::Join("`n", $lines.ToArray()))
} finally {
    if ($pres -ne $null) { try { $pres.Close() } catch {} }
    if ($ppt -ne $null) { try { if ($ppt.Presentations.Count -eq 0) { $ppt.Quit() } } catch {} }
    if ($pres -ne $null) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($pres) | Out-Null 2>$null }
    if ($ppt -ne $null) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ppt) | Out-Null 2>$null }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
