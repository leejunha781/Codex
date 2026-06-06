$ErrorActionPreference = 'Stop'

function Get-ProposalFolder {
    $matches = Get-ChildItem -LiteralPath 'D:\' -Recurse -Directory -ErrorAction SilentlyContinue |
        Where-Object { $_.FullName -like '*AVEVA - Marine Principal Technical Support & Consultant*Proposal' } |
        Select-Object -First 1
    if (-not $matches) {
        throw 'Could not locate AVEVA proposal folder under D:\.'
    }
    return $matches.FullName
}

function Get-PowerPointApplication {
    try {
        $app = New-Object -ComObject PowerPoint.Application
        return $app
    } catch {
        $exe = 'C:\Program Files\Microsoft Office\root\Office16\POWERPNT.EXE'
        if (-not (Test-Path -LiteralPath $exe)) {
            throw
        }
        Start-Process -FilePath $exe -WindowStyle Hidden
        Start-Sleep -Seconds 8
        return [Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application')
    }
}

$proposal = Get-ProposalFolder
$deckPath = Join-Path $proposal 'Future_Industrial_PLM_Meeting_Deck_EN_V13.pptx'
$defaultPath = Join-Path $proposal 'Future_Industrial_PLM_Meeting_Deck_EN.pptx'
$finalPath = Join-Path $proposal 'Future_Industrial_PLM_Meeting_Deck_EN_Final.pptx'
$pdfPath = Join-Path $proposal 'Future_Industrial_PLM_Meeting_Deck_EN_Final.pdf'
$scriptRoot = $PSScriptRoot
if ([string]::IsNullOrWhiteSpace($scriptRoot)) {
    $scriptRoot = Join-Path $env:USERPROFILE '.claude\plm_slide_work'
}
$reportPath = Join-Path $scriptRoot 'v13_slide2_reference_cards_verify.txt'

if (-not (Test-Path -LiteralPath $deckPath)) {
    throw "Deck not found: $deckPath"
}

$requiredShapeNames = @(
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
$oldShapeNames = @('GlossaryPointer', 'PresenterGuidePointer')

$ppt = $null
$pres = $null
$lines = New-Object System.Collections.Generic.List[string]

try {
    $ppt = Get-PowerPointApplication
    try {
        $ppt.Visible = [Microsoft.Office.Core.MsoTriState]::msoFalse
    } catch {
        # Some PowerPoint COM sessions reject hiding the application window.
    }
    $pres = $ppt.Presentations.Open($deckPath, [Microsoft.Office.Core.MsoTriState]::msoTrue, [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoFalse)

    $slideCount = $pres.Slides.Count
    $slide2 = $pres.Slides.Item(2)

    $shapeNames = @()
    $shapeTexts = @()
    for ($i = 1; $i -le $slide2.Shapes.Count; $i++) {
        $shape = $slide2.Shapes.Item($i)
        $shapeNames += $shape.Name
        $text = ''
        try {
            if ($shape.HasTextFrame -and $shape.TextFrame.HasText) {
                $text = ($shape.TextFrame.TextRange.Text -replace '[\r\n]+', ' ').Trim()
            }
        } catch {
            $text = ''
        }
        if ($text.Length -gt 0) {
            $shapeTexts += ('{0}: {1}' -f $shape.Name, $text)
        }
    }

    $missing = @($requiredShapeNames | Where-Object { $shapeNames -notcontains $_ })
    $oldRemaining = @($oldShapeNames | Where-Object { $shapeNames -contains $_ })
    $looseTextRemaining = @($shapeTexts | Where-Object { $_ -match 'Abbreviations\s*(->|→)' -or $_ -match 'Meeting procedure guide\s*(->|→)' })

    $lines.Add(('Deck: {0}' -f $deckPath))
    $lines.Add(('Slide count: {0}' -f $slideCount))
    $lines.Add(('Required V13 reference shapes present: {0}' -f (($missing.Count -eq 0).ToString())))
    $lines.Add(('Old pointer shapes absent: {0}' -f (($oldRemaining.Count -eq 0).ToString())))
    $lines.Add(('Loose old pointer text absent: {0}' -f (($looseTextRemaining.Count -eq 0).ToString())))
    $lines.Add(('Required missing count: {0}' -f $missing.Count))
    $lines.Add(('Old remaining count: {0}' -f $oldRemaining.Count))
    $lines.Add(('Loose text remaining count: {0}' -f $looseTextRemaining.Count))

    foreach ($path in @($deckPath, $defaultPath, $finalPath, $pdfPath)) {
        if (Test-Path -LiteralPath $path) {
            $item = Get-Item -LiteralPath $path
            $lines.Add(('File: {0} | {1} bytes | {2:yyyy-MM-dd HH:mm:ss}' -f $item.FullName, $item.Length, $item.LastWriteTime))
        } else {
            $lines.Add(('Missing file: {0}' -f $path))
        }
    }

    $lines.Add('Slide 2 reference card text:')
    foreach ($entry in $shapeTexts) {
        if ($entry -match 'MEETING PROCEDURE|Guide\s+p\.3|ABBREVIATIONS|Glossary\s+p\.39') {
            $lines.Add(('  {0}' -f $entry))
        }
    }

    if ($slideCount -ne 41 -or $missing.Count -ne 0 -or $oldRemaining.Count -ne 0 -or $looseTextRemaining.Count -ne 0) {
        $lines | Set-Content -LiteralPath $reportPath -Encoding UTF8
        throw ('Verification failed. See {0}' -f $reportPath)
    }

    $lines | Set-Content -LiteralPath $reportPath -Encoding UTF8
    $lines | ForEach-Object { Write-Output $_ }
} finally {
    if ($pres -ne $null) {
        $pres.Close()
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($pres)
    }
    if ($ppt -ne $null) {
        if ($ppt.Presentations.Count -eq 0) {
            $ppt.Quit()
        }
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($ppt)
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
