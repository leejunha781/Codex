$ErrorActionPreference = 'Stop'

function Get-ProposalFolder {
    $matches = Get-ChildItem -LiteralPath 'D:\' -Recurse -Directory -ErrorAction SilentlyContinue |
        Where-Object { $_.FullName -like '*AVEVA - Marine Principal Technical Support & Consultant*Proposal' } |
        Select-Object -First 1
    if (-not $matches) { throw 'Could not locate proposal folder.' }
    return $matches.FullName
}

function Get-PowerPointApplication {
    try {
        return New-Object -ComObject PowerPoint.Application
    } catch {
        $exe = 'C:\Program Files\Microsoft Office\root\Office16\POWERPNT.EXE'
        Start-Process -FilePath $exe -WindowStyle Hidden
        Start-Sleep -Seconds 8
        return [Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application')
    }
}

$proposal = Get-ProposalFolder
$v15 = Join-Path $proposal 'Future_Industrial_PLM_Meeting_Deck_EN_V15.pptx'
$default = Join-Path $proposal 'Future_Industrial_PLM_Meeting_Deck_EN.pptx'
$final = Join-Path $proposal 'Future_Industrial_PLM_Meeting_Deck_EN_Final.pptx'
$pdf = Join-Path $proposal 'Future_Industrial_PLM_Meeting_Deck_EN_Final.pdf'
$reportPath = Join-Path (Join-Path $env:USERPROFILE '.claude\plm_slide_work') 'v15_slide2_agenda_1_9_verify.txt'

$required = @()
for ($i = 1; $i -le 9; $i++) {
    $required += ('AgendaRowV15_{0:00}' -f $i)
    $required += ('AgendaNumberCircleV15_{0:00}' -f $i)
    $required += ('AgendaNumberTextV15_{0:00}' -f $i)
    $required += ('AgendaTitleV15_{0:00}' -f $i)
    $required += ('AgendaRangeChipV15_{0:00}' -f $i)
}
$oldReferenceNames = @(
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

$ppt = $null
$pres = $null
$lines = New-Object System.Collections.Generic.List[string]

try {
    $ppt = Get-PowerPointApplication
    try { $ppt.Visible = [Microsoft.Office.Core.MsoTriState]::msoFalse } catch {}
    $pres = $ppt.Presentations.Open($v15, [Microsoft.Office.Core.MsoTriState]::msoTrue, [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoFalse)
    $slide = $pres.Slides.Item(2)

    $shapeNames = @()
    $text = New-Object System.Collections.Generic.List[string]
    for ($i = 1; $i -le $slide.Shapes.Count; $i++) {
        $shape = $slide.Shapes.Item($i)
        $shapeNames += $shape.Name
        try {
            if ($shape.HasTextFrame -and $shape.TextFrame.HasText) {
                $text.Add(($shape.TextFrame.TextRange.Text -replace '[\r\n]+', ' ').Trim())
            }
        } catch {}
    }
    $textBlob = ($text -join ' ')

    $missing = @($required | Where-Object { $shapeNames -notcontains $_ })
    $oldRemaining = @($oldReferenceNames | Where-Object { $shapeNames -contains $_ })
    $hasNumbers = $true
    for ($i = 1; $i -le 9; $i++) {
        if ($text -notcontains ([string]$i)) { $hasNumbers = $false }
    }
    $hasIntegratedRows = $textBlob -match 'MEETING PROCEDURE' -and $textBlob -match 'ABBREVIATIONS / GLOSSARY' -and $textBlob -match 'p\.3–5' -and $textBlob -match 'p\.39–41'

    $lines.Add(('Deck: {0}' -f $v15))
    $lines.Add(('Slide count: {0}' -f $pres.Slides.Count))
    $lines.Add(('Required V15 agenda shapes present: {0}' -f (($missing.Count -eq 0).ToString())))
    $lines.Add(('Old floating reference card shapes absent: {0}' -f (($oldRemaining.Count -eq 0).ToString())))
    $lines.Add(('Numbers 1-9 present: {0}' -f $hasNumbers.ToString()))
    $lines.Add(('Integrated procedure/glossary rows present: {0}' -f $hasIntegratedRows.ToString()))
    $lines.Add(('Missing count: {0}' -f $missing.Count))
    $lines.Add(('Old remaining count: {0}' -f $oldRemaining.Count))

    foreach ($path in @($v15, $default, $final, $pdf)) {
        if (Test-Path -LiteralPath $path) {
            $item = Get-Item -LiteralPath $path
            $lines.Add(('File: {0} | {1} bytes | {2:yyyy-MM-dd HH:mm:ss}' -f $item.FullName, $item.Length, $item.LastWriteTime))
        } else {
            $lines.Add(('Missing file: {0}' -f $path))
        }
    }

    $lines | Set-Content -LiteralPath $reportPath -Encoding UTF8
    $lines | ForEach-Object { Write-Output $_ }

    if ($pres.Slides.Count -ne 41 -or $missing.Count -ne 0 -or $oldRemaining.Count -ne 0 -or -not $hasNumbers -or -not $hasIntegratedRows) {
        throw ('Verification failed. See {0}' -f $reportPath)
    }
} finally {
    if ($pres -ne $null) {
        $pres.Close()
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($pres)
    }
    if ($ppt -ne $null) {
        if ($ppt.Presentations.Count -eq 0) { $ppt.Quit() }
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($ppt)
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
