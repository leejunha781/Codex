$ErrorActionPreference = 'Stop'

$src = 'D:\이력서\AVEVA - Marine Principal Technical Support & Consultant – PLM SME, Busan\Proposal\Future_Industrial_PLM_Meeting_Deck_EN_V6.pptx'
$defaultAlias = 'D:\이력서\AVEVA - Marine Principal Technical Support & Consultant – PLM SME, Busan\Proposal\Future_Industrial_PLM_Meeting_Deck_EN.pptx'
$backupDir = 'D:\이력서\AVEVA - Marine Principal Technical Support & Consultant – PLM SME, Busan\Proposal\_backup_20260606_slide13_14_readability'

New-Item -ItemType Directory -Path $backupDir -Force | Out-Null
$timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$backupPath = Join-Path $backupDir ("Future_Industrial_PLM_Meeting_Deck_EN_V6_BEFORE_SLIDE13_14_READABILITY_{0}.pptx" -f $timestamp)
Copy-Item -LiteralPath $src -Destination $backupPath -Force

function Get-ShapeByName {
    param($Slide, [string]$Name)
    foreach ($shape in @($Slide.Shapes)) {
        if ($shape.Name -eq $Name) { return $shape }
    }
    throw "Shape not found on slide $($Slide.SlideIndex): $Name"
}

function Set-AllTextSize {
    param($Shape, [single]$Size)
    if ($Shape.HasTextFrame -and $Shape.TextFrame.HasText) {
        $Shape.TextFrame.TextRange.Font.Size = $Size
    }
}

function Set-TextShape {
    param(
        $Shape,
        [single]$Left,
        [single]$Top,
        [single]$Width,
        [single]$Height,
        [single]$FontSize
    )
    $Shape.Left = $Left
    $Shape.Top = $Top
    $Shape.Width = $Width
    $Shape.Height = $Height
    Set-AllTextSize -Shape $Shape -Size $FontSize
}

$pp = $null
$pres = $null

try {
    $pp = New-Object -ComObject PowerPoint.Application
    $pp.Visible = -1
    $pres = $pp.Presentations.Open($src, $false, $false, $false)

    $slide13 = $pres.Slides.Item(13)
    $slide14 = $pres.Slides.Item(14)

    # Slide 13: widen the two explanatory cards and enlarge their text;
    # keep the visual but trim its footprint so the content reads at meeting distance.
    $img13 = Get-ShapeByName $slide13 'Picture 96'
    $img13.Left = 594
    $img13.Top = 170
    $img13.Width = 330
    $img13.Height = 185.6

    $card13a = Get-ShapeByName $slide13 'Rounded Rectangle 53'
    $card13b = Get-ShapeByName $slide13 'Rounded Rectangle 61'
    $card13a.Left = 24
    $card13a.Top = 200
    $card13a.Width = 265
    $card13a.Height = 152
    $card13b.Left = 296
    $card13b.Top = 200
    $card13b.Width = 265
    $card13b.Height = 152

    Set-TextShape (Get-ShapeByName $slide13 'TextBox 54') 34 208 246 16 13.2
    Set-TextShape (Get-ShapeByName $slide13 'TextBox 62') 306 208 244 16 13.2

    $box13a = Get-ShapeByName $slide13 'Rounded Rectangle 55'
    $box13a.Left = 40
    $box13a.Top = 236
    $box13a.Width = 88
    $box13a.Height = 52
    Set-TextShape (Get-ShapeByName $slide13 'TextBox 56') 46 245 76 32 11.1

    $chip13a = Get-ShapeByName $slide13 'Rounded Rectangle 57'
    $chip13a.Left = 44
    $chip13a.Top = 302
    $chip13a.Width = 80
    $chip13a.Height = 15
    Set-TextShape (Get-ShapeByName $slide13 'TextBox 58') 46 305 76 8 8.5

    Set-TextShape (Get-ShapeByName $slide13 'TextBox 59') 136 235 136 86 9
    Set-TextShape (Get-ShapeByName $slide13 'TextBox 60') 136 325 136 27 9.3

    $box13b = Get-ShapeByName $slide13 'Rounded Rectangle 63'
    $box13b.Left = 308
    $box13b.Top = 236
    $box13b.Width = 84
    $box13b.Height = 52
    Set-TextShape (Get-ShapeByName $slide13 'TextBox 75') 317 278 66 9 8.4

    $chip13b = Get-ShapeByName $slide13 'Rounded Rectangle 76'
    $chip13b.Left = 309
    $chip13b.Top = 302
    $chip13b.Width = 78
    $chip13b.Height = 15
    Set-TextShape (Get-ShapeByName $slide13 'TextBox 77') 311 305 72 8 8.5

    Set-TextShape (Get-ShapeByName $slide13 'TextBox 78') 402 235 139 86 9
    Set-TextShape (Get-ShapeByName $slide13 'TextBox 79') 402 325 139 27 9.3

    $principle13 = Get-ShapeByName $slide13 'Rounded Rectangle 80'
    $principle13.Left = 24
    $principle13.Top = 366
    $principle13.Width = 536
    $principle13.Height = 29
    Set-TextShape (Get-ShapeByName $slide13 'TextBox 81') 36 374 73 10 9.4
    Set-TextShape (Get-ShapeByName $slide13 'TextBox 82') 114 371 426 14 9.8
    Set-TextShape (Get-ShapeByName $slide13 'TextBox 83') 28 398 526 8.5 7.2

    # Make the top 4-step strip slightly taller and easier to read.
    $stepBoxes13 = @('Rounded Rectangle 30','Rounded Rectangle 36','Rounded Rectangle 42','Rounded Rectangle 48')
    $stepNums13 = @('TextBox 32','TextBox 38','TextBox 44','TextBox 50')
    $stepTitles13 = @('TextBox 33','TextBox 39','TextBox 45','TextBox 51')
    $stepSubs13 = @('TextBox 34','TextBox 40','TextBox 46','TextBox 52')
    $stepLefts13 = @(24, 158, 292, 426)
    for ($i = 0; $i -lt 4; $i++) {
        $shape = Get-ShapeByName $slide13 $stepBoxes13[$i]
        $shape.Left = $stepLefts13[$i]
        $shape.Top = 156
        $shape.Width = 122
        $shape.Height = 38
        Set-TextShape (Get-ShapeByName $slide13 $stepNums13[$i]) ($stepLefts13[$i] + 6) 166 15 11 10.5
        Set-TextShape (Get-ShapeByName $slide13 $stepTitles13[$i]) ($stepLefts13[$i] + 24) 163 90 12 10.8
        Set-TextShape (Get-ShapeByName $slide13 $stepSubs13[$i]) ($stepLefts13[$i] + 24) 176 92 9 8.2
    }

    # Slide 14: enlarge the tier roadmap and both explanatory panels.
    $img14 = Get-ShapeByName $slide14 'Picture 1'
    $img14.Left = 592
    $img14.Top = 160
    $img14.Width = 332
    $img14.Height = 186.8

    $tierBoxes14 = @('Rounded Rectangle 30','Rounded Rectangle 37','Rounded Rectangle 44','Rounded Rectangle 51')
    $tierSmallTags14 = @('TextBox 31','TextBox 38','TextBox 45','TextBox 52')
    $tierTitles14 = @('TextBox 32','TextBox 39','TextBox 46','TextBox 53')
    $tierSubs14 = @('TextBox 33','TextBox 40','TextBox 47','TextBox 54')
    $tierPills14 = @('Rounded Rectangle 34','Rounded Rectangle 41','Rounded Rectangle 48','Rounded Rectangle 55')
    $tierPillTexts14 = @('TextBox 35','TextBox 42','TextBox 49','TextBox 56')
    $tierLefts14 = @(24, 156, 288, 420)
    for ($i = 0; $i -lt 4; $i++) {
        $shape = Get-ShapeByName $slide14 $tierBoxes14[$i]
        $shape.Left = $tierLefts14[$i]
        $shape.Top = 160
        $shape.Width = 127
        $shape.Height = 48
        Set-TextShape (Get-ShapeByName $slide14 $tierSmallTags14[$i]) ($tierLefts14[$i] + 7) 166 32 9 8.6
        Set-TextShape (Get-ShapeByName $slide14 $tierTitles14[$i]) ($tierLefts14[$i] + 7) 176 115 14 13.2
        Set-TextShape (Get-ShapeByName $slide14 $tierSubs14[$i]) ($tierLefts14[$i] + 7) 190 115 10 8.7
        $pill = Get-ShapeByName $slide14 $tierPills14[$i]
        $pill.Left = $tierLefts14[$i] + 8
        $pill.Top = 202
        $pill.Width = 62
        $pill.Height = 12
        Set-TextShape (Get-ShapeByName $slide14 $tierPillTexts14[$i]) ($tierLefts14[$i] + 10) 205 58 8.7 8.5
    }

    $panel14a = Get-ShapeByName $slide14 'Rounded Rectangle 57'
    $panel14b = Get-ShapeByName $slide14 'Rounded Rectangle 75'
    $panel14a.Left = 24
    $panel14a.Top = 224
    $panel14a.Width = 264
    $panel14a.Height = 122
    $panel14b.Left = 294
    $panel14b.Top = 224
    $panel14b.Width = 264
    $panel14b.Height = 122
    Set-TextShape (Get-ShapeByName $slide14 'TextBox 58') 35 231 236 13 12.8
    Set-TextShape (Get-ShapeByName $slide14 'TextBox 76') 303 231 230 13 12.8

    $stackRects14 = @('Rounded Rectangle 59','Rounded Rectangle 63','Rounded Rectangle 67','Rounded Rectangle 71')
    $stackTopTexts14 = @('TextBox 60','TextBox 64','TextBox 68','TextBox 72')
    $stackBottomTexts14 = @('TextBox 61','TextBox 65','TextBox 69','TextBox 73')
    $stackLefts14 = @(40, 102, 164, 226)
    for ($i = 0; $i -lt 4; $i++) {
        $shape = Get-ShapeByName $slide14 $stackRects14[$i]
        $shape.Left = $stackLefts14[$i]
        $shape.Top = 252
        $shape.Width = 50
        $shape.Height = 34
        Set-TextShape (Get-ShapeByName $slide14 $stackTopTexts14[$i]) ($stackLefts14[$i] + 2) 258 46 18 7.7
        Set-TextShape (Get-ShapeByName $slide14 $stackBottomTexts14[$i]) ($stackLefts14[$i] + 2) 276 46 7 6.9
    }

    Set-TextShape (Get-ShapeByName $slide14 'TextBox 74') 36 296 244 44 8.5

    $reasonTitleBoxes = @('TextBox 79','TextBox 83','TextBox 87','TextBox 91')
    $reasonBodyBoxes = @('TextBox 80','TextBox 84','TextBox 88','TextBox 92')
    foreach ($name in $reasonTitleBoxes) {
        $shape = Get-ShapeByName $slide14 $name
        Set-AllTextSize -Shape $shape -Size 10.2
    }
    foreach ($name in $reasonBodyBoxes) {
        $shape = Get-ShapeByName $slide14 $name
        Set-AllTextSize -Shape $shape -Size 8.4
    }

    $interfaceBar14 = Get-ShapeByName $slide14 'Rounded Rectangle 93'
    $interfaceBar14.Left = 24
    $interfaceBar14.Top = 360
    $interfaceBar14.Width = 536
    $interfaceBar14.Height = 22
    Set-TextShape (Get-ShapeByName $slide14 'TextBox 94') 36 365 514 10 9.1
    Set-TextShape (Get-ShapeByName $slide14 'TextBox 95') 36 384 514 8.5 7.4

    $pres.Save()
    Copy-Item -LiteralPath $src -Destination $defaultAlias -Force
}
finally {
    if ($pres) { $pres.Close() }
    if ($pp) { $pp.Quit() }
}

Write-Output "Edited: $src"
Write-Output "Backup: $backupPath"
Write-Output "Default alias refreshed: $defaultAlias"
