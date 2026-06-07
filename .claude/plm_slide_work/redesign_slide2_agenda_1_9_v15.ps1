$ErrorActionPreference = 'Stop'

function Get-ProposalFolder {
    $matches = Get-ChildItem -LiteralPath 'D:\' -Recurse -Directory -ErrorAction SilentlyContinue |
        Where-Object { $_.FullName -like '*AVEVA - Marine Principal Technical Support & Consultant*Proposal' } |
        Select-Object -First 1
    if (-not $matches) { throw 'Could not locate AVEVA proposal folder under D:\.' }
    return $matches.FullName
}

function Get-PowerPointApplication {
    try {
        return New-Object -ComObject PowerPoint.Application
    } catch {
        $exe = 'C:\Program Files\Microsoft Office\root\Office16\POWERPNT.EXE'
        if (-not (Test-Path -LiteralPath $exe)) { throw }
        Start-Process -FilePath $exe -WindowStyle Hidden
        Start-Sleep -Seconds 8
        return [Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application')
    }
}

function Rgb($r, $g, $b) {
    return [int]($r + ($g * 256) + ($b * 65536))
}

function Add-TextBox($slide, $name, $left, $top, $width, $height, $text, $fontSize, $color, $bold) {
    $shape = $slide.Shapes.AddTextbox(1, [single]$left, [single]$top, [single]$width, [single]$height)
    $shape.Name = $name
    $shape.TextFrame.TextRange.Text = $text
    $shape.TextFrame.MarginLeft = [single]0
    $shape.TextFrame.MarginRight = [single]0
    $shape.TextFrame.MarginTop = [single]0
    $shape.TextFrame.MarginBottom = [single]0
    $shape.TextFrame.WordWrap = -1
    $shape.TextFrame.TextRange.Font.Name = 'Aptos'
    $shape.TextFrame.TextRange.Font.Size = [single]$fontSize
    $shape.TextFrame.TextRange.Font.Color.RGB = $color
    $shape.TextFrame.TextRange.Font.Bold = [int]$(if ($bold) { -1 } else { 0 })
    return $shape
}

function Add-AgendaRow($slide, $idx, $title, $desc, $range, $color, $top) {
    $left = 48
    $width = 864
    $height = 46
    $white = Rgb 238 246 255
    $muted = Rgb 184 200 218
    $rowFill = Rgb 17 38 61
    $chipFill = Rgb 12 31 50

    $row = $slide.Shapes.AddShape(5, [single]$left, [single]$top, [single]$width, [single]$height)
    $row.Name = ('AgendaRowV15_{0:00}' -f $idx)
    $row.Fill.ForeColor.RGB = $rowFill
    $row.Fill.Transparency = [single]0.04
    $row.Line.ForeColor.RGB = $color
    $row.Line.Weight = [single]1.0
    $row.Line.Transparency = [single]0.0
    try { $row.Adjustments.Item(1) = [single]0.08 } catch {}

    $dot = $slide.Shapes.AddShape(9, [single]60, [single]($top + 9), [single]28, [single]28)
    $dot.Name = ('AgendaNumberCircleV15_{0:00}' -f $idx)
    $dot.Fill.ForeColor.RGB = $color
    $dot.Line.Visible = 0

    $num = Add-TextBox $slide ('AgendaNumberTextV15_{0:00}' -f $idx) 60 ($top + 9) 28 29.1 ([string]$idx) 13 $white $true
    $num.TextFrame.TextRange.ParagraphFormat.Alignment = 2

    Add-TextBox $slide ('AgendaTitleV15_{0:00}' -f $idx) 102 ($top + 7) 260 15 $title 9.6 $color $true | Out-Null
    Add-TextBox $slide ('AgendaDescV15_{0:00}' -f $idx) 102 ($top + 23) 640 14 $desc 7.7 $muted $false | Out-Null

    $chip = $slide.Shapes.AddShape(5, [single]858, [single]($top + 14), [single]50, [single]17)
    $chip.Name = ('AgendaRangeChipV15_{0:00}' -f $idx)
    $chip.Fill.ForeColor.RGB = $chipFill
    $chip.Line.ForeColor.RGB = $color
    $chip.Line.Weight = [single]0.8
    try { $chip.Adjustments.Item(1) = [single]0.15 } catch {}

    $chip.TextFrame.TextRange.Text = $range
    $chip.TextFrame.MarginLeft = [single]0
    $chip.TextFrame.MarginRight = [single]0
    $chip.TextFrame.MarginTop = [single]0
    $chip.TextFrame.MarginBottom = [single]0
    $chip.TextFrame.TextRange.Font.Name = 'Aptos'
    $chip.TextFrame.TextRange.Font.Size = [single]8.6
    $chip.TextFrame.TextRange.Font.Bold = [int]-1
    $chip.TextFrame.TextRange.Font.Color.RGB = $color
    $chip.TextFrame.TextRange.ParagraphFormat.Alignment = 2
}

$proposal = Get-ProposalFolder
$source = Join-Path $proposal 'Future_Industrial_PLM_Meeting_Deck_EN_Final.pptx'
$v15 = Join-Path $proposal 'Future_Industrial_PLM_Meeting_Deck_EN_V15.pptx'
$default = Join-Path $proposal 'Future_Industrial_PLM_Meeting_Deck_EN.pptx'
$final = Join-Path $proposal 'Future_Industrial_PLM_Meeting_Deck_EN_Final.pptx'
$pdf = Join-Path $proposal 'Future_Industrial_PLM_Meeting_Deck_EN_Final.pdf'
$stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$backupDir = Join-Path $proposal ("_backup_20260607_slide2_agenda_1_9_v15_{0}" -f $stamp)
$reportPath = Join-Path (Join-Path $env:USERPROFILE '.claude\plm_slide_work') 'slide2_agenda_1_9_v15_report.txt'

New-Item -ItemType Directory -Path $backupDir -Force | Out-Null
foreach ($path in @($source, $default, $final, $pdf, $v15)) {
    if (Test-Path -LiteralPath $path) {
        Copy-Item -LiteralPath $path -Destination (Join-Path $backupDir (Split-Path $path -Leaf)) -Force
    }
}
Copy-Item -LiteralPath $source -Destination $v15 -Force

$ppt = $null
$pres = $null
$lines = New-Object System.Collections.Generic.List[string]

try {
    $ppt = Get-PowerPointApplication
    try { $ppt.Visible = [Microsoft.Office.Core.MsoTriState]::msoFalse } catch {}
    $pres = $ppt.Presentations.Open($v15, [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoFalse)
    $slide = $pres.Slides.Item(2)

    $deleted = 0
    for ($i = $slide.Shapes.Count; $i -ge 1; $i--) {
        $shape = $slide.Shapes.Item($i)
        $text = ''
        try {
            if ($shape.HasTextFrame -and $shape.TextFrame.HasText) {
                $text = ($shape.TextFrame.TextRange.Text -replace '[\r\n]+', ' ').Trim()
            }
        } catch {}

        $remove = $false
        if ($shape.Name -match '^(MeetingNavRange[0-9]+|ReferenceToolbarPanelV13|ProcedureGuide|GlossaryGuide|AgendaRowV15|AgendaNumber|AgendaTitleV15|AgendaDescV15|AgendaRangeChipV15)') { $remove = $true }
        if ($shape.Left -ge 45 -and $shape.Left -le 915 -and $shape.Top -ge 95 -and $shape.Top -le 510) { $remove = $true }
        if ($text -match 'Procedure rule:|CURRENT STATE|PUBLIC SIGNALS|WINNING IDEAS|PROPOSAL \+ DELIVERY|FUTURE-STATE COMPARE|REVIEW|CLOSE|MEETING PROCEDURE|ABBREVIATIONS|Guide\s+p\.3|Glossary\s+p\.39') { $remove = $true }

        # Keep slide number, title, subtitle, divider and page brand.
        if ($shape.Name -in @('TextBox 11','TextBox 13','TextBox 14','Rectangle 15','TextBox 53')) { $remove = $false }

        if ($remove) {
            $shape.Delete()
            $deleted++
        }
    }

    $cyan = Rgb 42 190 230
    $blue = Rgb 78 139 239
    $magenta = Rgb 228 72 185
    $green = Rgb 76 204 154
    $orange = Rgb 255 139 44
    $yellow = Rgb 244 188 49
    $purple = Rgb 155 119 226
    $teal = Rgb 43 218 164

    $items = @(
        @{ Title='MEETING PROCEDURE'; Desc='Align the room on decision flow before facts, ideas and delivery plan'; Range='p.3–5'; Color=$teal },
        @{ Title='CURRENT STATE'; Desc='Competitor vs AVEVA comparison, PLM position and executable baseline'; Range='p.6–8'; Color=$cyan },
        @{ Title='PUBLIC SIGNALS'; Desc='Official product and roadmap signals; confirmed versus directional'; Range='p.9'; Color=$blue },
        @{ Title='WINNING IDEAS'; Desc='Meeting thesis + win-above NAPA / Siemens / Dassault patterns'; Range='p.10–18'; Color=$magenta },
        @{ Title='PROPOSAL + DELIVERY'; Desc='Architecture, workflows and pilot scope - then WBS and KPI gates'; Range='p.19–32'; Color=$green },
        @{ Title='FUTURE-STATE COMPARE'; Desc='Re-score AVEVA after the proposed ideas are applied'; Range='p.33'; Color=$orange },
        @{ Title='REVIEW'; Desc='Ownership, risks, evidence and assumptions requiring agreement'; Range='p.34–36'; Color=$yellow },
        @{ Title='CLOSE'; Desc='Approve the next increment and the assurance pilot'; Range='p.37–38'; Color=$purple },
        @{ Title='ABBREVIATIONS / GLOSSARY'; Desc='Reference appendix for acronyms and architecture terms'; Range='p.39–41'; Color=$cyan }
    )

    $top = 86
    $gap = 2
    for ($idx = 1; $idx -le $items.Count; $idx++) {
        $item = $items[$idx - 1]
        Add-AgendaRow $slide $idx $item.Title $item.Desc $item.Range $item.Color ($top + (($idx - 1) * (46 + $gap)))
    }

    $revisionUpdated = 0
    $cover = $pres.Slides.Item(1)
    for ($i = 1; $i -le $cover.Shapes.Count; $i++) {
        $shape = $cover.Shapes.Item($i)
        try {
            if ($shape.HasTextFrame -and $shape.TextFrame.HasText) {
                if ($shape.TextFrame.TextRange.Text -match 'Rev\. V\d+') {
                    $shape.TextFrame.TextRange.Text = 'Rev. V15 · 2026-06-07'
                    $revisionUpdated++
                }
            }
        } catch {}
    }

    $pres.Save()
    $pres.SaveCopyAs($default)
    $pres.SaveCopyAs($final)
    if (Test-Path -LiteralPath $pdf) { Remove-Item -LiteralPath $pdf -Force }
    $pres.SaveAs($pdf, 32)

    $lines.Add(('Source: {0}' -f $source))
    $lines.Add(('V15: {0}' -f $v15))
    $lines.Add(('Default: {0}' -f $default))
    $lines.Add(('Final: {0}' -f $final))
    $lines.Add(('PDF: {0}' -f $pdf))
    $lines.Add(('Backup: {0}' -f $backupDir))
    $lines.Add(('Slide count: {0}' -f $pres.Slides.Count))
    $lines.Add(('Slide 2 deleted prior agenda/reference shapes: {0}' -f $deleted))
    $lines.Add(('Cover revision replacements: {0}' -f $revisionUpdated))
    $lines | Set-Content -LiteralPath $reportPath -Encoding UTF8
    $lines | ForEach-Object { Write-Output $_ }
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
