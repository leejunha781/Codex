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
$v14 = Join-Path $proposal 'Future_Industrial_PLM_Meeting_Deck_EN_V14.pptx'
$default = Join-Path $proposal 'Future_Industrial_PLM_Meeting_Deck_EN.pptx'
$final = Join-Path $proposal 'Future_Industrial_PLM_Meeting_Deck_EN_Final.pptx'
$pdf = Join-Path $proposal 'Future_Industrial_PLM_Meeting_Deck_EN_Final.pdf'
$reportPath = Join-Path (Join-Path $env:USERPROFILE '.claude\plm_slide_work') 'v14_slide7_stack_cards_verify.txt'

$required = @(
    'StackPositionPanelV14',
    'StackPositionTitleV14',
    'StackCardV14_1_AVEVA',
    'StackCardV14_2_Siemens',
    'StackCardV14_3_Dassault',
    'StackCardV14_4_Hexagon',
    'StackCardV14_5_PTC',
    'StackPositionInsightV14'
)
$oldNames = @('TextBox 35','Oval 36','TextBox 37','TextBox 38','TextBox 39','Oval 40','TextBox 41','TextBox 42','TextBox 43','Oval 44','TextBox 45','TextBox 46','TextBox 47','Oval 48','TextBox 49','TextBox 50','TextBox 51','Oval 52','TextBox 53','TextBox 54','TextBox 55')

$ppt = $null
$pres = $null
$lines = New-Object System.Collections.Generic.List[string]

try {
    $ppt = Get-PowerPointApplication
    try { $ppt.Visible = [Microsoft.Office.Core.MsoTriState]::msoFalse } catch {}
    $pres = $ppt.Presentations.Open($v14, [Microsoft.Office.Core.MsoTriState]::msoTrue, [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoFalse)
    $slide = $pres.Slides.Item(7)

    $shapeNames = @()
    $textValues = @()
    for ($i = 1; $i -le $slide.Shapes.Count; $i++) {
        $shape = $slide.Shapes.Item($i)
        $shapeNames += $shape.Name
        try {
            if ($shape.HasTextFrame -and $shape.TextFrame.HasText) {
                $textValues += (($shape.TextFrame.TextRange.Text -replace '[\r\n]+', ' ').Trim())
            }
        } catch {}
    }

    $missing = @($required | Where-Object { $shapeNames -notcontains $_ })
    $oldRemaining = @($oldNames | Where-Object { $shapeNames -contains $_ })
    $hasKeyText = ($textValues -join ' ') -match 'Who owns which layer' -and ($textValues -join ' ') -match 'Engineering → Operations bridge' -and ($textValues -join ' ') -match 'PLM \+ IoT digital thread'

    $lines.Add(('Deck: {0}' -f $v14))
    $lines.Add(('Slide count: {0}' -f $pres.Slides.Count))
    $lines.Add(('Required V14 stack-card shapes present: {0}' -f (($missing.Count -eq 0).ToString())))
    $lines.Add(('Old slide 7 stack-list shape names absent: {0}' -f (($oldRemaining.Count -eq 0).ToString())))
    $lines.Add(('Key redesigned slide 7 text present: {0}' -f $hasKeyText.ToString()))
    $lines.Add(('Missing count: {0}' -f $missing.Count))
    $lines.Add(('Old remaining count: {0}' -f $oldRemaining.Count))

    foreach ($path in @($v14, $default, $final, $pdf)) {
        if (Test-Path -LiteralPath $path) {
            $item = Get-Item -LiteralPath $path
            $lines.Add(('File: {0} | {1} bytes | {2:yyyy-MM-dd HH:mm:ss}' -f $item.FullName, $item.Length, $item.LastWriteTime))
        } else {
            $lines.Add(('Missing file: {0}' -f $path))
        }
    }

    $lines | Set-Content -LiteralPath $reportPath -Encoding UTF8
    $lines | ForEach-Object { Write-Output $_ }

    if ($pres.Slides.Count -ne 41 -or $missing.Count -ne 0 -or $oldRemaining.Count -ne 0 -or -not $hasKeyText) {
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
