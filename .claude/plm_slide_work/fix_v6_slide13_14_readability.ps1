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
    $img13.Left = 580
    $img13.Top = 172
    $img13.Width = 344
    $img13.Height = 193.5

    $card13a = Get-ShapeByName $slide13 'Rounded Rectangle 53'
    $card13b = Get-ShapeByName $slide13 'Rounded Rectangle 61'
    $card13a.Left = 24
    $card13a.Top = 200
    $card13a.Width = 257
    $card13a.Height = 150
    $card13b.Left = 288
    $card13b.Top = 200
    $card13b.Width = 257
    $card13b.Height = 150

    Set-TextShape (Get-ShapeByName $slide13 'TextBox 54') 34 208 235 16 13.5
    Set-TextShape (Get-ShapeByName $slide13 'TextBox 62') 298 208 231 16 13.5

    $box13a = Get-ShapeByName $slide13 'Rounded Rectangle 55'
    $box13a.Left = 40
    $box13a.Top = 236
    $box13a.Width = 92
    $box13a.Height = 52
    Set-TextShape (Get-ShapeByName $slide13 'TextBox 56') 48 245 76 32 11.5

    $chip13a = Get-ShapeByName $slide13 'Rounded Rectangle 57'
    $chip13a.Left = 45
    $chip13a.Top = 300
    $chip13a.Width = 82
    $chip13a.Height = 15
    Set-TextShape (Get-ShapeByName $slide13 'TextBox 58') 47 303.5 78 9 9

    Set-TextShape (Get-ShapeByName $slide13 'TextBox 59') 145 236 112 73 10
    Set-TextShape (Get-ShapeByName $slide13 'TextBox 60') 145 315 113 20 10.5

    $box13b = Get-ShapeByName $slide13 'Rounded Rectangle 63'
    $box13b.Left = 303
    $box13b.Top = 236
    $box13b.Width = 84
    $box13b.Height = 52
    Set-TextShape (Get-ShapeByName $slide13 'TextBox 75') 313 278 62 9 8.7

    $chip13b = Get-ShapeByName $slide13 'Rounded Rectangle 76'
    $chip13b.Left = 306
    $chip13b.Top = 300
    $chip13b.Width = 78
    $chip13b.Height = 15
    Set-TextShape (Get-ShapeByName $slide13 'TextBox 77') 308 303.5 74 9 9

    Set-TextShape (Get-ShapeByName $slide13 'TextBox 78') 398 236 123 73 10
    Set-TextShape (Get-ShapeByName $slide13 'TextBox 79') 398 315 122 20 10.5

    $principle13 = Get-ShapeByName $slide13 'Rounded Rectangle 80'
    $principle13.Left = 24
    $principle13.Top = 362
    $principle13.Width = 520
    $principle13.Height = 29
    Set-TextShape (Get-ShapeByName $slide13 'TextBox 81') 36 370 73 10 9.6
    Set-TextShape (Get-ShapeByName $slide13 'TextBox 82') 110 367.5 416 14 10.2
    Set-TextShape (Get-ShapeByName $slide13 'TextBox 83') 28 397 510 8.5 7.4

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
    $img14.Left = 578
    $img14.Top = 160
    $img14.Width = 346
    $img14.Height = 194.6

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
        $shape.Width = 124
        $shape.Height = 48
        Set-TextShape (Get-ShapeByName $slide14 $tierSmallTags14[$i]) ($tierLefts14[$i] + 7) 166 32 9 8.6
        Set-TextShape (Get-ShapeByName $slide14 $tierTitles14[$i]) ($tierLefts14[$i] + 7) 176 112 14 13.3
        Set-TextShape (Get-ShapeByName $slide14 $tierSubs14[$i]) ($tierLefts14[$i] + 7) 190 112 10 8.9
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
    $panel14a.Width = 255
    $panel14a.Height = 120
    $panel14b.Left = 286
    $panel14b.Top = 224
    $panel14b.Width = 255
    $panel14b.Height = 120
    Set-TextShape (Get-ShapeByName $slide14 'TextBox 58') 35 231 228 13 13
    Set-TextShape (Get-ShapeByName $slide14 'TextBox 76') 297 231 228 13 13

    $stackRects14 = @('Rounded Rectangle 59','Rounded Rectangle 63','Rounded Rectangle 67','Rounded Rectangle 71')
    $stackTopTexts14 = @('TextBox 60','TextBox 64','TextBox 68','TextBox 72')
    $stackBottomTexts14 = @('TextBox 61','TextBox 65','TextBox 69','TextBox 73')
    $stackLefts14 = @(38, 96, 154, 212)
    for ($i = 0; $i -lt 4; $i++) {
        $shape = Get-ShapeByName $slide14 $stackRects14[$i]
        $shape.Left = $stackLefts14[$i]
        $shape.Top = 252
        $shape.Width = 48
        $shape.Height = 34
        Set-TextShape (Get-ShapeByName $slide14 $stackTopTexts14[$i]) ($stackLefts14[$i] + 2) 258 44 18 8.7
        Set-TextShape (Get-ShapeByName $slide14 $stackBottomTexts14[$i]) ($stackLefts14[$i] + 2) 276 44 7 7.5
    }

    Set-TextShape (Get-ShapeByName $slide14 'TextBox 74') 36 296 234 43 9.2

    $reasonTitleBoxes = @('TextBox 79','TextBox 83','TextBox 87','TextBox 91')
    $reasonBodyBoxes = @('TextBox 80','TextBox 84','TextBox 88','TextBox 92')
    foreach ($name in $reasonTitleBoxes) {
        $shape = Get-ShapeByName $slide14 $name
        Set-AllTextSize -Shape $shape -Size 10.4
    }
    foreach ($name in $reasonBodyBoxes) {
        $shape = Get-ShapeByName $slide14 $name
        Set-AllTextSize -Shape $shape -Size 8.6
    }

    $interfaceBar14 = Get-ShapeByName $slide14 'Rounded Rectangle 93'
    $interfaceBar14.Left = 24
    $interfaceBar14.Top = 356
    $interfaceBar14.Width = 517
    $interfaceBar14.Height = 24
    Set-TextShape (Get-ShapeByName $slide14 'TextBox 94') 36 362 492 10 9.7
    Set-TextShape (Get-ShapeByName $slide14 'TextBox 95') 36 383 492 8.5 7.6

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
