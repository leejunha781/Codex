$ErrorActionPreference = 'Stop'

$proposalDir = 'D:\이력서\AVEVA - Marine Principal Technical Support & Consultant – PLM SME, Busan\Proposal'
$sourceDeck = Join-Path $proposalDir 'Future_Industrial_PLM_Meeting_Deck_EN_Final.pptx'
$v13Deck = Join-Path $proposalDir 'Future_Industrial_PLM_Meeting_Deck_EN_V13.pptx'
$defaultDeck = Join-Path $proposalDir 'Future_Industrial_PLM_Meeting_Deck_EN.pptx'
$finalDeck = Join-Path $proposalDir 'Future_Industrial_PLM_Meeting_Deck_EN_Final.pptx'
$finalPdf = Join-Path $proposalDir 'Future_Industrial_PLM_Meeting_Deck_EN_Final.pdf'
$workDir = 'C:\Users\namma\.claude\plm_slide_work'
$timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$backupDir = Join-Path $proposalDir "_backup_20260607_slide2_reference_cards_v13_$timestamp"
$reportPath = Join-Path $workDir 'future_direction_meeting_v13_slide2_reference_cards_report.txt'

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

function Backup-IfExists {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][string]$BackupDir
    )
    if (Test-Path -LiteralPath $Path) {
        Copy-Item -LiteralPath $Path -Destination (Join-Path $BackupDir ([System.IO.Path]::GetFileName($Path))) -Force
    }
}

function Rgb {
    param([int]$R, [int]$G, [int]$B)
    return ($R + ($G * 256) + ($B * 65536))
}

function Set-TextContainingInShapes {
    param(
        $Shapes,
        [string]$Needle,
        [string]$NewText,
        [ref]$Count
    )
    foreach ($shape in $Shapes) {
        try {
            if ($shape.Type -eq 6) {
                Set-TextContainingInShapes -Shapes $shape.GroupItems -Needle $Needle -NewText $NewText -Count $Count
            } elseif ($shape.HasTextFrame -and $shape.TextFrame.HasText) {
                if ($shape.TextFrame.TextRange.Text -like "*$Needle*") {
                    $shape.TextFrame.TextRange.Text = $NewText
                    $Count.Value = $Count.Value + 1
                }
            }
        } catch {}
    }
}

function Remove-ShapeByNameOrText {
    param(
        $Slide,
        [string[]]$Names,
        [string[]]$TextNeedles,
        [ref]$Count
    )
    for ($i = $Slide.Shapes.Count; $i -ge 1; $i--) {
        $shape = $Slide.Shapes.Item($i)
        $delete = $false
        foreach ($name in $Names) {
            if ($shape.Name -eq $name) { $delete = $true }
        }
        if (-not $delete) {
            try {
                if ($shape.HasTextFrame -and $shape.TextFrame.HasText) {
                    $txt = $shape.TextFrame.TextRange.Text
                    foreach ($needle in $TextNeedles) {
                        if ($txt -like "*$needle*") { $delete = $true }
                    }
                }
            } catch {}
        }
        if ($delete) {
            $shape.Delete()
            $Count.Value = $Count.Value + 1
        }
    }
}

function Add-NavCard {
    param(
        $Slide,
        [string]$NamePrefix,
        [float]$Left,
        [float]$Top,
        [float]$Width,
        [float]$Height,
        [string]$Label,
        [string]$RangeText,
        [int]$AccentColor,
        [int]$FillColor,
        [int]$WhiteColor,
        [int]$MutedColor
    )
    $card = $Slide.Shapes.AddShape(5, $Left, $Top, $Width, $Height)
    $card.Name = "${NamePrefix}CardV13"
    $card.Fill.Visible = -1
    $card.Fill.ForeColor.RGB = $FillColor
    $card.Fill.Transparency = [single]0.04
    $card.Line.Visible = -1
    $card.Line.ForeColor.RGB = $AccentColor
    $card.Line.Weight = [single]1.35
    try { $card.Adjustments.Item(1) = 0.18 } catch {}

    $bar = $Slide.Shapes.AddShape(1, $Left, $Top, 4, $Height)
    $bar.Name = "${NamePrefix}AccentBarV13"
    $bar.Fill.ForeColor.RGB = $AccentColor
    $bar.Line.Visible = 0

    $labelBox = $Slide.Shapes.AddTextbox(1, $Left + 12, $Top + 7, $Width - 24, 12)
    $labelBox.Name = "${NamePrefix}LabelV13"
    $labelBox.TextFrame.TextRange.Text = $Label
    $labelBox.TextFrame.TextRange.Font.Name = 'Aptos Display'
    $labelBox.TextFrame.TextRange.Font.Size = [single]6.8
    $labelBox.TextFrame.TextRange.Font.Bold = [int]-1
    $labelBox.TextFrame.TextRange.Font.Color.RGB = $MutedColor
    $labelBox.TextFrame.MarginLeft = 0
    $labelBox.TextFrame.MarginRight = 0
    $labelBox.TextFrame.MarginTop = 0
    $labelBox.TextFrame.MarginBottom = 0

    $rangeBox = $Slide.Shapes.AddTextbox(1, $Left + 12, $Top + 19, $Width - 24, 14)
    $rangeBox.Name = "${NamePrefix}RangeV13"
    $rangeBox.TextFrame.TextRange.Text = $RangeText
    $rangeBox.TextFrame.TextRange.Font.Name = 'Aptos Display'
    $rangeBox.TextFrame.TextRange.Font.Size = [single]8.7
    $rangeBox.TextFrame.TextRange.Font.Bold = [int]-1
    $rangeBox.TextFrame.TextRange.Font.Color.RGB = $WhiteColor
    $rangeBox.TextFrame.MarginLeft = 0
    $rangeBox.TextFrame.MarginRight = 0
    $rangeBox.TextFrame.MarginTop = 0
    $rangeBox.TextFrame.MarginBottom = 0
}

if (-not (Test-Path -LiteralPath $sourceDeck)) { throw "Missing source deck: $sourceDeck" }

New-Item -ItemType Directory -Path $backupDir -Force | Out-Null
Backup-IfExists -Path $sourceDeck -BackupDir $backupDir
Backup-IfExists -Path $v13Deck -BackupDir $backupDir
Backup-IfExists -Path $defaultDeck -BackupDir $backupDir
Backup-IfExists -Path $finalDeck -BackupDir $backupDir
Backup-IfExists -Path $finalPdf -BackupDir $backupDir
Copy-Item -LiteralPath $sourceDeck -Destination $v13Deck -Force

$ppt = $null
$pres = $null
$removed = 0
$revCount = 0
try {
    $ppt = Get-PowerPointApplication
    $pres = $ppt.Presentations.Open($v13Deck, $false, $false, $false)

    Set-TextContainingInShapes -Shapes $pres.Slides.Item(1).Shapes -Needle 'Rev. V12' -NewText 'Rev. V13 · 2026-06-07' -Count ([ref]$revCount)

    $slide = $pres.Slides.Item(2)
    Remove-ShapeByNameOrText -Slide $slide `
        -Names @('GlossaryPointer', 'PresenterGuidePointer', 'ProcedureGuidePointer', 'ReferenceToolbarPanelV13') `
        -TextNeedles @('Abbreviations', 'Glossary appendix', 'Meeting procedure guide') `
        -Count ([ref]$removed)

    $panelFill = Rgb 8 29 47
    $cyan = Rgb 34 214 230
    $teal = Rgb 50 214 156
    $white = Rgb 245 250 255
    $muted = Rgb 176 199 213

    $rail = $slide.Shapes.AddShape(5, 526, 51, 392, 42)
    $rail.Name = 'ReferenceToolbarPanelV13'
    $rail.Fill.Visible = -1
    $rail.Fill.ForeColor.RGB = Rgb 5 20 35
    $rail.Fill.Transparency = [single]0.08
    $rail.Line.Visible = -1
    $rail.Line.ForeColor.RGB = Rgb 36 76 100
    $rail.Line.Weight = [single]0.8
    try { $rail.Adjustments.Item(1) = 0.18 } catch {}

    Add-NavCard -Slide $slide -NamePrefix 'ProcedureGuide' -Left 538 -Top 56 -Width 176 -Height 32 `
        -Label 'MEETING PROCEDURE' -RangeText 'Guide  p.3–5' -AccentColor $teal `
        -FillColor $panelFill -WhiteColor $white -MutedColor $muted
    Add-NavCard -Slide $slide -NamePrefix 'GlossaryGuide' -Left 724 -Top 56 -Width 184 -Height 32 `
        -Label 'ABBREVIATIONS' -RangeText 'Glossary  p.39–41' -AccentColor $cyan `
        -FillColor $panelFill -WhiteColor $white -MutedColor $muted

    $pres.Save()
    $pres.Close()
    $pres = $null
} finally {
    if ($pres -ne $null) { try { $pres.Close() } catch {} }
    if ($ppt -ne $null) { try { if ($ppt.Presentations.Count -eq 0) { $ppt.Quit() } } catch {} }
    if ($pres -ne $null) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($pres) | Out-Null 2>$null }
    if ($ppt -ne $null) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ppt) | Out-Null 2>$null }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

Copy-Item -LiteralPath $v13Deck -Destination $defaultDeck -Force
Copy-Item -LiteralPath $v13Deck -Destination $finalDeck -Force

$ppt = $null
$pres = $null
try {
    $ppt = Get-PowerPointApplication
    $pres = $ppt.Presentations.Open($finalDeck, $true, $false, $false)
    $pres.SaveAs($finalPdf, 32)
    $pres.Close()
    $pres = $null
} finally {
    if ($pres -ne $null) { try { $pres.Close() } catch {} }
    if ($ppt -ne $null) { try { if ($ppt.Presentations.Count -eq 0) { $ppt.Quit() } } catch {} }
    if ($pres -ne $null) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($pres) | Out-Null 2>$null }
    if ($ppt -ne $null) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ppt) | Out-Null 2>$null }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

$lines = New-Object System.Collections.Generic.List[string]
[void]$lines.Add('Future Industrial PLM V13 - slide 2 reference-card redesign')
[void]$lines.Add("Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss zzz')")
[void]$lines.Add('Scope: fix loose top reference text on slide 2 by moving content into designed navigation cards.')
[void]$lines.Add("Backup folder: $backupDir")
[void]$lines.Add('')
[void]$lines.Add("Removed loose pointer/text shapes: $removed")
[void]$lines.Add("Cover revision replacements: $revCount")
[void]$lines.Add('Changes:')
[void]$lines.Add('- Removed loose GlossaryPointer / PresenterGuidePointer text from slide 2.')
[void]$lines.Add('- Added right-side reference toolbar with two rounded navigation cards.')
[void]$lines.Add('- Card 1: MEETING PROCEDURE / Guide p.3–5.')
[void]$lines.Add('- Card 2: ABBREVIATIONS / Glossary p.39–41.')
[void]$lines.Add('- Cover revision updated to Rev. V13 · 2026-06-07.')
[void]$lines.Add('')
[void]$lines.Add('Saved files:')
foreach ($path in @($v13Deck, $defaultDeck, $finalDeck, $finalPdf)) {
    $item = Get-Item -LiteralPath $path
    [void]$lines.Add(("{0}`t{1}`t{2}" -f $item.FullName, $item.Length, $item.LastWriteTime.ToString('yyyy-MM-dd HH:mm:ss')))
}
[System.IO.File]::WriteAllLines($reportPath, $lines.ToArray(), [System.Text.Encoding]::UTF8)
Write-Host ([string]::Join("`n", $lines.ToArray()))
