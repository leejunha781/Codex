$ErrorActionPreference = 'Stop'

$proposal = 'D:\이력서\AVEVA - Marine Principal Technical Support & Consultant – PLM SME, Busan\Proposal'
$source = Join-Path $proposal 'Future_Industrial_PLM_Meeting_Deck_EN_V8.pptx'
$v9 = Join-Path $proposal 'Future_Industrial_PLM_Meeting_Deck_EN_V9.pptx'
$defaultDeck = Join-Path $proposal 'Future_Industrial_PLM_Meeting_Deck_EN.pptx'
$finalDeck = Join-Path $proposal 'Future_Industrial_PLM_Meeting_Deck_EN_Final.pptx'
$finalPdf = Join-Path $proposal 'Future_Industrial_PLM_Meeting_Deck_EN_Final.pdf'
$workDir = 'C:\Users\namma\.claude\plm_slide_work'
$renderDir = Join-Path $workDir 'future_direction_meeting_v9_slides3_5_visibility_render'
$reportPath = Join-Path $workDir 'future_direction_meeting_v9_slides3_5_visibility_report.txt'
$stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$backupDir = Join-Path $proposal ("_backup_20260607_slides3_5_visibility_v9_" + $stamp)

if (-not (Test-Path -LiteralPath $source)) { throw "Missing source deck: $source" }
if (Test-Path -LiteralPath $renderDir) { Remove-Item -LiteralPath $renderDir -Recurse -Force }
New-Item -ItemType Directory -Path $renderDir | Out-Null
New-Item -ItemType Directory -Path $backupDir | Out-Null

foreach ($p in @($source, $v9, $defaultDeck, $finalDeck, $finalPdf)) {
    if (Test-Path -LiteralPath $p) {
        Copy-Item -LiteralPath $p -Destination (Join-Path $backupDir (Split-Path $p -Leaf)) -Force
    }
}
Copy-Item -LiteralPath $source -Destination $v9 -Force

function Rgb {
    param([int]$R, [int]$G, [int]$B)
    return [int]($R + ($G -shl 8) + ($B -shl 16))
}

$C = @{
    Navy = (Rgb 7 18 32)
    Navy2 = (Rgb 10 30 49)
    Panel = (Rgb 15 40 62)
    Panel2 = (Rgb 19 49 75)
    Cyan = (Rgb 28 202 214)
    Cyan2 = (Rgb 56 214 238)
    Teal = (Rgb 55 211 153)
    Amber = (Rgb 255 176 73)
    Blue = (Rgb 87 153 255)
    Magenta = (Rgb 233 73 190)
    Red = (Rgb 255 105 105)
    White = (Rgb 255 255 255)
    Muted = (Rgb 198 213 228)
    Dim = (Rgb 137 160 182)
}

function New-PowerPoint {
    try {
        return New-Object -ComObject PowerPoint.Application
    } catch {
        $exe = 'C:\Program Files\Microsoft Office\root\Office16\POWERPNT.EXE'
        if (-not (Test-Path -LiteralPath $exe)) { $exe = 'POWERPNT.EXE' }
        Start-Process -FilePath $exe -ArgumentList '/automation' -WindowStyle Hidden
        Start-Sleep -Seconds 5
        return New-Object -ComObject PowerPoint.Application
    }
}

function Clear-Slide {
    param([object]$Slide)
    for ($i = $Slide.Shapes.Count; $i -ge 1; $i--) {
        $Slide.Shapes.Item($i).Delete()
    }
    $Slide.FollowMasterBackground = $false
    $Slide.Background.Fill.ForeColor.RGB = $C.Navy
}

function Add-Text {
    param(
        [object]$Slide,
        [double]$Left,
        [double]$Top,
        [double]$Width,
        [double]$Height,
        [string]$Text,
        [single]$Size,
        [int]$Color,
        [int]$Bold = 0,
        [int]$Align = 1,
        [string]$Name = ''
    )
    $shape = $Slide.Shapes.AddTextbox(1, $Left, $Top, $Width, $Height)
    if ($Name.Length -gt 0) { $shape.Name = $Name }
    $shape.TextFrame.TextRange.Text = $Text
    $shape.TextFrame.TextRange.Font.Name = 'Aptos'
    $shape.TextFrame.TextRange.Font.Size = $Size
    $shape.TextFrame.TextRange.Font.Color.RGB = $Color
    $shape.TextFrame.TextRange.Font.Bold = $Bold
    $shape.TextFrame.TextRange.ParagraphFormat.Alignment = $Align
    $shape.TextFrame.WordWrap = -1
    $shape.TextFrame.MarginLeft = 4
    $shape.TextFrame.MarginRight = 4
    $shape.TextFrame.MarginTop = 2
    $shape.TextFrame.MarginBottom = 2
    return $shape
}

function Add-Base {
    param([object]$Slide, [int]$No, [string]$Title, [string]$Subtitle)
    Clear-Slide -Slide $Slide

    # Subtle technical background: editable line mesh, not a raster image.
    for ($i = 0; $i -lt 8; $i++) {
        $x1 = 40 + ($i * 115)
        $x2 = $x1 + 170
        $line = $Slide.Shapes.AddLine($x1, 500, $x2, 120)
        $line.Line.ForeColor.RGB = $C.Panel2
        $line.Line.Transparency = 0.55
        $line.Line.Weight = 0.8
    }
    for ($i = 0; $i -lt 9; $i++) {
        $dot = $Slide.Shapes.AddShape(9, 70 + ($i * 95), 430 - (($i % 3) * 35), 5, 5)
        $dot.Fill.ForeColor.RGB = $C.Cyan
        $dot.Fill.Transparency = 0.45
        $dot.Line.Visible = 0
    }

    Add-Text -Slide $Slide -Left 36 -Top 18 -Width 780 -Height 34 -Text $Title -Size ([single]22) -Color $C.White -Bold 1 -Name 'Title' | Out-Null
    Add-Text -Slide $Slide -Left 38 -Top 55 -Width 820 -Height 22 -Text $Subtitle -Size ([single]10.2) -Color $C.Muted -Bold 0 -Name 'Subtitle' | Out-Null
    $divider = $Slide.Shapes.AddShape(1, 36, 84, 858, 2.2)
    $divider.Fill.ForeColor.RGB = $C.Cyan
    $divider.Line.Visible = 0
    Add-Text -Slide $Slide -Left 36 -Top 508 -Width 400 -Height 16 -Text 'Future Industrial PLM / Meeting Procedure' -Size ([single]7.8) -Color $C.Dim -Bold 0 | Out-Null
    Add-Text -Slide $Slide -Left 875 -Top 508 -Width 35 -Height 16 -Text ([string]$No) -Size ([single]7.8) -Color $C.Muted -Bold 0 -Align 3 | Out-Null
}

function Add-Card {
    param(
        [object]$Slide,
        [double]$Left,
        [double]$Top,
        [double]$Width,
        [double]$Height,
        [string]$Header,
        [string]$Body,
        [int]$Accent,
        [string]$Name = ''
    )
    $card = $Slide.Shapes.AddShape(5, $Left, $Top, $Width, $Height)
    if ($Name.Length -gt 0) { $card.Name = $Name }
    $card.Fill.ForeColor.RGB = $C.Panel
    $card.Fill.Transparency = 0.02
    $card.Line.ForeColor.RGB = $Accent
    $card.Line.Weight = 1.3
    $bar = $Slide.Shapes.AddShape(1, $Left, $Top, 6, $Height)
    $bar.Fill.ForeColor.RGB = $Accent
    $bar.Line.Visible = 0
    Add-Text -Slide $Slide -Left ($Left + 16) -Top ($Top + 10) -Width ($Width - 28) -Height 22 -Text $Header -Size ([single]12.2) -Color $Accent -Bold 1 | Out-Null
    Add-Text -Slide $Slide -Left ($Left + 16) -Top ($Top + 38) -Width ($Width - 30) -Height ($Height - 44) -Text $Body -Size ([single]9.8) -Color $C.White -Bold 0 | Out-Null
}

function Add-NumberCircle {
    param([object]$Slide, [double]$Left, [double]$Top, [double]$Size, [string]$No, [int]$Fill)
    $circle = $Slide.Shapes.AddShape(9, $Left, $Top, $Size, $Size)
    $circle.Fill.ForeColor.RGB = $Fill
    $circle.Line.Visible = 0
    Add-Text -Slide $Slide -Left $Left -Top ($Top + ($Size * 0.20)) -Width $Size -Height ($Size * 0.55) -Text $No -Size ([single]18) -Color $C.Navy -Bold 1 -Align 2 | Out-Null
}

function Add-Arrow {
    param([object]$Slide, [double]$X1, [double]$Y1, [double]$X2, [double]$Y2, [int]$Color)
    $line = $Slide.Shapes.AddLine($X1, $Y1, $X2, $Y2)
    $line.Line.ForeColor.RGB = $Color
    $line.Line.Weight = 2.25
    $line.Line.EndArrowheadStyle = 3
}

function Build-Slide3 {
    param([object]$Slide)
    Add-Base -Slide $Slide -No 3 -Title 'Meeting Procedure - decision path' -Subtitle 'The room first agrees how to decide, then follows the evidence from gap to pilot approval.'

    $steps = @(
        @{No='1'; H='START WITH THE GAP'; B='No vendor owns the complete change-to-assurance-to-operations loop.'; C=$C.Cyan},
        @{No='2'; H='MAKE THE MOVE'; B='Position AVEVA Marine as the open decision control plane above specialist tools.'; C=$C.Teal},
        @{No='3'; H='PROVE FEASIBILITY'; B='Windows authoring + Linux control plane + Local Agent / Plugin adapters.'; C=$C.Amber},
        @{No='4'; H='CLOSE THE PILOT'; B='Approve Phase 2 configuration core + one bounded assurance pilot.'; C=$C.Cyan2}
    )

    $x = 52
    for ($i = 0; $i -lt $steps.Count; $i++) {
        $s = $steps[$i]
        $left = $x + ($i * 217)
        Add-NumberCircle -Slide $Slide -Left ($left + 74) -Top 112 -Size 54 -No $s.No -Fill $s.C
        Add-Card -Slide $Slide -Left $left -Top 178 -Width 185 -Height 155 -Header $s.H -Body $s.B -Accent $s.C -Name ("ProcedureStep" + $s.No)
        if ($i -lt 3) {
            Add-Arrow -Slide $Slide -X1 ($left + 190) -Y1 254 -X2 ($left + 213) -Y2 254 -Color $C.Cyan
        }
    }

    Add-Card -Slide $Slide -Left 58 -Top 374 -Width 548 -Height 72 -Header 'PRESENTER RULE' -Body 'Do not sell this as a tool-replacement story. Sell it as a governed, auditable, operations-aware decision loop that makes existing tools more valuable.' -Accent $C.Cyan
    Add-Card -Slide $Slide -Left 632 -Top 374 -Width 252 -Height 72 -Header 'DECISION OUTPUT' -Body 'Product direction, pilot scope, owners and first evidence scenario.' -Accent $C.Amber
}

function Build-Slide4 {
    param([object]$Slide)
    Add-Base -Slide $Slide -No 4 -Title 'What development teams must hear' -Subtitle 'Translate the future direction into buildable interfaces, validation gates and explicit anti-patterns.'

    Add-Card -Slide $Slide -Left 44 -Top 118 -Width 260 -Height 260 -Header 'BUILD' -Body "OpenAPI contracts for objects, BOM, effectivity, baseline, change impact and assurance evidence.`r`n`r`nLocal Agent / Plugin callbacks from Windows authoring tools.`r`n`r`nAI Gate provider interface: rules first, RAG next, CONNECT later." -Accent $C.Cyan -Name 'BuildCard'
    Add-Card -Slide $Slide -Left 333 -Top 118 -Width 260 -Height 260 -Header 'VALIDATE' -Body "Seed truth cases for hull, block, compartment, weight, loading case and affected scenarios.`r`n`r`nE2E tests for evaluate -> review -> promote and rollback.`r`n`r`nKPI gates for contract coverage, accuracy, provenance, response time and security." -Accent $C.Teal -Name 'ValidateCard'
    Add-Card -Slide $Slide -Left 622 -Top 118 -Width 260 -Height 260 -Header 'AVOID' -Body "No direct Linux hosting assumption for E3D / Marine / Hull.`r`n`r`nNo big-bang replacement of certified solvers.`r`n`r`nNo automatic AI release authority.`r`n`r`nNo unbounded process customization without approval gates." -Accent $C.Red -Name 'AvoidCard'

    # Interface spine.
    $spine = $Slide.Shapes.AddShape(5, 70, 408, 785, 58)
    $spine.Fill.ForeColor.RGB = $C.Navy2
    $spine.Fill.Transparency = 0.02
    $spine.Line.ForeColor.RGB = $C.Cyan
    $spine.Line.Weight = 1.2
    Add-Text -Slide $Slide -Left 92 -Top 419 -Width 130 -Height 26 -Text 'Windows Agent' -Size ([single]10.5) -Color $C.White -Bold 1 -Align 2 | Out-Null
    Add-Text -Slide $Slide -Left 270 -Top 419 -Width 110 -Height 26 -Text 'OpenAPI' -Size ([single]10.5) -Color $C.White -Bold 1 -Align 2 | Out-Null
    Add-Text -Slide $Slide -Left 428 -Top 419 -Width 140 -Height 26 -Text 'SQL Evidence' -Size ([single]10.5) -Color $C.White -Bold 1 -Align 2 | Out-Null
    Add-Text -Slide $Slide -Left 620 -Top 419 -Width 100 -Height 26 -Text 'CI Gate' -Size ([single]10.5) -Color $C.White -Bold 1 -Align 2 | Out-Null
    Add-Arrow -Slide $Slide -X1 224 -Y1 435 -X2 268 -Y2 435 -Color $C.Cyan
    Add-Arrow -Slide $Slide -X1 382 -Y1 435 -X2 426 -Y2 435 -Color $C.Cyan
    Add-Arrow -Slide $Slide -X1 570 -Y1 435 -X2 618 -Y2 435 -Color $C.Cyan
    Add-Text -Slide $Slide -Left 742 -Top 419 -Width 95 -Height 26 -Text 'Trusted state' -Size ([single]10.5) -Color $C.Amber -Bold 1 -Align 2 | Out-Null
}

function Build-Slide5 {
    param([object]$Slide)
    Add-Base -Slide $Slide -No 5 -Title 'What planning teams must hear' -Subtitle 'Make the product story easy to explain: market position, customer value, roadmap discipline and decision ask.'

    Add-Card -Slide $Slide -Left 50 -Top 112 -Width 380 -Height 118 -Header 'PRODUCT POSITION' -Body 'AVEVA owns the control plane above specialist tools. The customer keeps tool choice; AVEVA owns lifecycle decision, evidence and approval orchestration.' -Accent $C.Cyan
    Add-Card -Slide $Slide -Left 494 -Top 112 -Width 380 -Height 118 -Header 'CUSTOMER VALUE' -Body 'Fewer blind handovers, faster change impact review, visible assurance evidence, explainable approvals and replayable design-to-operations history.' -Accent $C.Teal
    Add-Card -Slide $Slide -Left 50 -Top 260 -Width 380 -Height 118 -Header 'ROADMAP DISCIPLINE' -Body 'A 26-week gate model prevents wishful scope. Each capability must pass measurable acceptance before promotion.' -Accent $C.Amber
    Add-Card -Slide $Slide -Left 494 -Top 260 -Width 380 -Height 118 -Header 'DECISION REQUEST' -Body 'Approve bounded Phase 2 + Continuous Naval Assurance, agree owners, KPI thresholds, assumption boundaries and pilot scenario.' -Accent $C.Cyan2

    # Central value loop, drawn as editable shapes.
    $hub = $Slide.Shapes.AddShape(9, 408, 202, 106, 106)
    $hub.Fill.ForeColor.RGB = $C.Navy2
    $hub.Line.ForeColor.RGB = $C.Cyan
    $hub.Line.Weight = 2
    Add-Text -Slide $Slide -Left 421 -Top 224 -Width 80 -Height 42 -Text "AVEVA`r`nCONTROL`r`nPLANE" -Size ([single]10) -Color $C.White -Bold 1 -Align 2 | Out-Null
    $ring = $Slide.Shapes.AddShape(9, 395, 189, 132, 132)
    $ring.Fill.Visible = 0
    $ring.Line.ForeColor.RGB = $C.Cyan
    $ring.Line.Transparency = 0.55
    $ring.Line.Weight = 1.1

    $success = $Slide.Shapes.AddShape(5, 95, 406, 735, 76)
    $success.Fill.ForeColor.RGB = $C.Panel
    $success.Fill.Transparency = 0.02
    $success.Line.ForeColor.RGB = $C.Teal
    $success.Line.Weight = 1.25
    $successBar = $Slide.Shapes.AddShape(1, 95, 406, 6, 76)
    $successBar.Fill.ForeColor.RGB = $C.Teal
    $successBar.Line.Visible = 0
    Add-Text -Slide $Slide -Left 122 -Top 418 -Width 690 -Height 18 -Text 'PLANNING SUCCESS STATEMENT' -Size ([single]11.8) -Color $C.Teal -Bold 1 | Out-Null
    Add-Text -Slide $Slide -Left 122 -Top 448 -Width 690 -Height 24 -Text 'AVEVA Marine becomes the trusted system for seeing, approving and learning from every major shipbuilding change.' -Size ([single]10.2) -Color $C.White -Bold 0 | Out-Null
}

$pp = New-PowerPoint
$presentation = $null
try {
    $presentation = $pp.Presentations.Open($v9, $false, $false, $false)
    if ($presentation.Slides.Count -ne 41) { throw "Expected 41 slides, found $($presentation.Slides.Count)" }

    Build-Slide3 -Slide $presentation.Slides.Item(3)
    Build-Slide4 -Slide $presentation.Slides.Item(4)
    Build-Slide5 -Slide $presentation.Slides.Item(5)

    # Cover revision bump.
    $s1 = $presentation.Slides.Item(1)
    for ($i = 1; $i -le $s1.Shapes.Count; $i++) {
        $shape = $s1.Shapes.Item($i)
        try {
            if (($shape.HasTextFrame -eq -1) -and ($shape.TextFrame.HasText -eq -1)) {
                $txt = $shape.TextFrame.TextRange.Text
                if ($txt -like '*Rev. V8*') { $shape.TextFrame.TextRange.Text = $txt.Replace('Rev. V8', 'Rev. V9') }
            }
        } catch {}
    }

    $presentation.Save()
    $presentation.SaveAs($finalPdf, 32)
    for ($i = 1; $i -le $presentation.Slides.Count; $i++) {
        $png = Join-Path $renderDir ("slide{0:D2}.png" -f $i)
        $presentation.Slides.Item($i).Export($png, 'PNG', 1600, 900)
    }
    $presentation.Close()
    $presentation = $null
} finally {
    if ($presentation -ne $null) { $presentation.Close() }
    if ($pp.Presentations.Count -eq 0) { $pp.Quit() }
}

Copy-Item -LiteralPath $v9 -Destination $defaultDeck -Force
Copy-Item -LiteralPath $v9 -Destination $finalDeck -Force

$pngCount = (Get-ChildItem -LiteralPath $renderDir -Filter '*.png' | Measure-Object).Count
$files = @($v9, $defaultDeck, $finalDeck, $finalPdf) | ForEach-Object { Get-Item -LiteralPath $_ }

$report = New-Object System.Collections.Generic.List[string]
$report.Add('Future Industrial PLM V9 - redesigned slides 3-5 for visibility')
$report.Add(("Generated: {0:yyyy-MM-dd HH:mm:ss K}" -f (Get-Date)))
$report.Add('Scope: redesigned p.3 meeting procedure, p.4 development team message, p.5 planning team message.')
$report.Add('Image decision: no raster image generated; slides use editable PowerPoint diagrams/icons because they improve readability without adding file dependencies.')
$report.Add(("PNG render count: {0}" -f $pngCount))
$report.Add(("Render folder: {0}" -f $renderDir))
$report.Add(("Backup folder: {0}" -f $backupDir))
$report.Add('')
$report.Add('Saved files:')
foreach ($f in $files) {
    $report.Add(("{0}`t{1}`t{2:yyyy-MM-dd HH:mm:ss}" -f $f.FullName, $f.Length, $f.LastWriteTime))
}
[System.IO.File]::WriteAllLines($reportPath, $report, [System.Text.Encoding]::UTF8)
Get-Content -LiteralPath $reportPath
