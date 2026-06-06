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

function Add-RoundRect($slide, $name, $left, $top, $width, $height, $fill, $line, $lineWeight, $radiusHint) {
    $shape = $slide.Shapes.AddShape(5, [single]$left, [single]$top, [single]$width, [single]$height)
    $shape.Name = $name
    $shape.Fill.Visible = -1
    $shape.Fill.ForeColor.RGB = $fill
    $shape.Fill.Transparency = [single]0.08
    $shape.Line.Visible = -1
    $shape.Line.ForeColor.RGB = $line
    $shape.Line.Transparency = [single]0.18
    $shape.Line.Weight = [single]$lineWeight
    try { $shape.Adjustments.Item(1) = [single]$radiusHint } catch {}
    return $shape
}

function Add-Card($slide, $idx, $vendor, $stack, $position, $color, $top) {
    $x = 637
    $w = 287
    $h = 50
    $cardFill = Rgb 16 33 52
    $cardLine = Rgb 53 92 125
    $white = Rgb 238 246 255
    $muted = Rgb 188 203 219
    $dim = Rgb 145 166 188

    $card = Add-RoundRect $slide ("StackCardV14_{0}_{1}" -f $idx, $vendor) $x $top $w $h $cardFill $cardLine 0.9 0.08

    $bar = $slide.Shapes.AddShape(1, [single]($x + 1), [single]($top + 4), [single]5, [single]($h - 8))
    $bar.Name = ("StackAccentV14_{0}_{1}" -f $idx, $vendor)
    $bar.Fill.ForeColor.RGB = $color
    $bar.Line.Visible = 0

    $dot = $slide.Shapes.AddShape(9, [single]($x + 14), [single]($top + 15), [single]16, [single]16)
    $dot.Name = ("StackDotV14_{0}_{1}" -f $idx, $vendor)
    $dot.Fill.ForeColor.RGB = $color
    $dot.Line.Visible = 0

    $vendorShape = Add-TextBox $slide ("StackVendorV14_{0}_{1}" -f $idx, $vendor) ($x + 39) ($top + 6) 96 16 $vendor 10.5 $white $true
    $stackShape = Add-TextBox $slide ("StackProductsV14_{0}_{1}" -f $idx, $vendor) ($x + 134) ($top + 7) 145 15 $stack 7.9 $muted $false
    $positionShape = Add-TextBox $slide ("StackPositionV14_{0}_{1}" -f $idx, $vendor) ($x + 39) ($top + 26) 238 16 $position 8.5 $dim $false

    # Shapes are added in visual stacking order: card, accents, then text.
}

$proposal = Get-ProposalFolder
$source = Join-Path $proposal 'Future_Industrial_PLM_Meeting_Deck_EN_Final.pptx'
$v14 = Join-Path $proposal 'Future_Industrial_PLM_Meeting_Deck_EN_V14.pptx'
$default = Join-Path $proposal 'Future_Industrial_PLM_Meeting_Deck_EN.pptx'
$final = Join-Path $proposal 'Future_Industrial_PLM_Meeting_Deck_EN_Final.pptx'
$pdf = Join-Path $proposal 'Future_Industrial_PLM_Meeting_Deck_EN_Final.pdf'
$stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$backupDir = Join-Path $proposal ("_backup_20260607_slide7_stack_cards_v14_{0}" -f $stamp)
$reportPath = Join-Path (Join-Path $env:USERPROFILE '.claude\plm_slide_work') 'slide7_stack_cards_v14_report.txt'

New-Item -ItemType Directory -Path $backupDir -Force | Out-Null
foreach ($path in @($source, $default, $final, $pdf, $v14)) {
    if (Test-Path -LiteralPath $path) {
        Copy-Item -LiteralPath $path -Destination (Join-Path $backupDir (Split-Path $path -Leaf)) -Force
    }
}
Copy-Item -LiteralPath $source -Destination $v14 -Force

$ppt = $null
$pres = $null
$lines = New-Object System.Collections.Generic.List[string]

try {
    $ppt = Get-PowerPointApplication
    try { $ppt.Visible = [Microsoft.Office.Core.MsoTriState]::msoFalse } catch {}
    $pres = $ppt.Presentations.Open($v14, [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoFalse)

    $slide = $pres.Slides.Item(7)
    $deleteNames = @(
        'TextBox 35','Oval 36','TextBox 37','TextBox 38','TextBox 39',
        'Oval 40','TextBox 41','TextBox 42','TextBox 43',
        'Oval 44','TextBox 45','TextBox 46','TextBox 47',
        'Oval 48','TextBox 49','TextBox 50','TextBox 51',
        'Oval 52','TextBox 53','TextBox 54','TextBox 55'
    )

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
        if ($deleteNames -contains $shape.Name) { $remove = $true }
        if ($shape.Left -ge 620 -and $shape.Top -ge 80 -and $shape.Top -le 390) {
            if ($text -match 'Stack & Position|E3D/Marine|Teamcenter|3DEXPERIENCE|Smart 3D|Windchill|PLM core backbone|digital thread') {
                $remove = $true
            }
        }
        if ($remove) {
            $shape.Delete()
            $deleted++
        }
    }

    $navy = Rgb 8 23 39
    $panelFill = Rgb 12 29 47
    $panelLine = Rgb 58 95 128
    $teal = Rgb 42 190 230
    $white = Rgb 238 246 255
    $muted = Rgb 170 190 209
    $green = Rgb 76 204 154
    $purple = Rgb 155 119 226
    $orange = Rgb 255 139 44
    $blueGrey = Rgb 157 183 207

    $panel = Add-RoundRect $slide 'StackPositionPanelV14' 625 88 310 363 $panelFill $panelLine 1.2 0.05
    $panel.Fill.Transparency = [single]0.0

    $title = Add-TextBox $slide 'StackPositionTitleV14' 641 101 230 20 'Stack & Position' 13.5 $white $true
    $sub = Add-TextBox $slide 'StackPositionSubtitleV14' 641 121 274 16 'Who owns which layer in the marine PLM battlefield' 7.8 $muted $false

    $rule = $slide.Shapes.AddShape(1, [single]641, [single]142, [single]270, [single]1.4)
    $rule.Name = 'StackPositionRuleV14'
    $rule.Fill.ForeColor.RGB = $teal
    $rule.Fill.Transparency = [single]0.15
    $rule.Line.Visible = 0

    Add-Card $slide 1 'AVEVA' 'E3D/Marine · AIM · PI · CONNECT' 'Engineering → Operations bridge' $teal 155
    Add-Card $slide 2 'Siemens' 'Teamcenter · NX · Opcenter' 'PLM core backbone' $blueGrey 211
    Add-Card $slide 3 'Dassault' '3DEXPERIENCE · CATIA · DELMIA' 'Experience / virtual-twin platform' $purple 267
    Add-Card $slide 4 'Hexagon' 'SDx · Smart 3D · j5 · Completions' 'Asset lifecycle & operations' $green 323
    Add-Card $slide 5 'PTC' 'Windchill · Codebeamer · ThingWorx' 'PLM + IoT digital thread' $orange 379

    $insight = Add-RoundRect $slide 'StackPositionInsightV14' 637 436 287 24 (Rgb 5 47 54) $green 0.8 0.08
    $insight.Fill.Transparency = [single]0.05
    $insightText = Add-TextBox $slide 'StackPositionInsightTextV14' 648 441 265 13 'Implication: AVEVA wins by connecting engineering truth to live operations.' 7.4 $white $false

    # Panel is created first so the card/table shapes remain on top.

    $revisionUpdated = 0
    $cover = $pres.Slides.Item(1)
    for ($i = 1; $i -le $cover.Shapes.Count; $i++) {
        $shape = $cover.Shapes.Item($i)
        try {
            if ($shape.HasTextFrame -and $shape.TextFrame.HasText) {
                $text = $shape.TextFrame.TextRange.Text
                if ($text -match 'Rev\. V\d+') {
                    $shape.TextFrame.TextRange.Text = 'Rev. V14 · 2026-06-07'
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
    $lines.Add(('V14: {0}' -f $v14))
    $lines.Add(('Default: {0}' -f $default))
    $lines.Add(('Final: {0}' -f $final))
    $lines.Add(('PDF: {0}' -f $pdf))
    $lines.Add(('Backup: {0}' -f $backupDir))
    $lines.Add(('Slide count: {0}' -f $pres.Slides.Count))
    $lines.Add(('Removed old slide 7 stack list shapes: {0}' -f $deleted))
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
