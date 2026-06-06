$ErrorActionPreference = 'Stop'

$proposal = 'D:\이력서\AVEVA - Marine Principal Technical Support & Consultant – PLM SME, Busan\Proposal'
$sourceDeck = Join-Path $proposal 'Future_Industrial_PLM_Meeting_Deck_EN_V6.pptx'
$v7Deck = Join-Path $proposal 'Future_Industrial_PLM_Meeting_Deck_EN_V7.pptx'
$defaultDeck = Join-Path $proposal 'Future_Industrial_PLM_Meeting_Deck_EN.pptx'
$finalDeck = Join-Path $proposal 'Future_Industrial_PLM_Meeting_Deck_EN_Final.pptx'
$finalPdf = Join-Path $proposal 'Future_Industrial_PLM_Meeting_Deck_EN_Final.pdf'
$briefMd = Join-Path $proposal 'Future_Industrial_PLM_Development_Planning_Meeting_Brief_EN.md'
$briefTxt = Join-Path $proposal 'Future_Industrial_PLM_Development_Planning_Meeting_Brief_EN.txt'
$workDir = 'C:\Users\namma\.claude\plm_slide_work'
$renderDir = Join-Path $workDir 'future_direction_meeting_v7_render'
$reportPath = Join-Path $workDir 'future_direction_meeting_v7_report.txt'
$stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$backupDir = Join-Path $proposal ("_backup_20260606_future_direction_v7_" + $stamp)

if (-not (Test-Path -LiteralPath $sourceDeck)) { throw "Source deck not found: $sourceDeck" }
if (-not (Test-Path -LiteralPath $workDir)) { New-Item -ItemType Directory -Path $workDir | Out-Null }
if (Test-Path -LiteralPath $renderDir) { Remove-Item -LiteralPath $renderDir -Recurse -Force }
New-Item -ItemType Directory -Path $renderDir | Out-Null
New-Item -ItemType Directory -Path $backupDir | Out-Null

foreach ($path in @($sourceDeck, $v7Deck, $defaultDeck, $finalDeck, $finalPdf, $briefMd, $briefTxt)) {
    if (Test-Path -LiteralPath $path) {
        Copy-Item -LiteralPath $path -Destination (Join-Path $backupDir (Split-Path $path -Leaf)) -Force
    }
}

Copy-Item -LiteralPath $sourceDeck -Destination $v7Deck -Force

function Rgb {
    param([int]$R, [int]$G, [int]$B)
    return [int]($R + ($G -shl 8) + ($B -shl 16))
}

$C = @{
    Navy = (Rgb 9 23 39)
    Panel = (Rgb 16 35 55)
    Panel2 = (Rgb 20 48 70)
    Cyan = (Rgb 27 202 214)
    Teal = (Rgb 51 207 153)
    Amber = (Rgb 255 176 73)
    White = (Rgb 255 255 255)
    Muted = (Rgb 193 210 226)
    Dim = (Rgb 137 160 182)
    Red = (Rgb 255 112 112)
}

function Add-Text {
    param(
        [object]$Slide,
        [double]$Left,
        [double]$Top,
        [double]$Width,
        [double]$Height,
        [string]$Text,
        [single]$Size = 10,
        [int]$Color,
        [int]$Bold = 0,
        [string]$Name = ''
    )
    $shape = $Slide.Shapes.AddTextbox(1, $Left, $Top, $Width, $Height)
    if ($Name.Length -gt 0) { $shape.Name = $Name }
    $shape.TextFrame.TextRange.Text = $Text
    $shape.TextFrame.TextRange.Font.Name = 'Aptos'
    $shape.TextFrame.TextRange.Font.Size = $Size
    $shape.TextFrame.TextRange.Font.Color.RGB = $Color
    $shape.TextFrame.TextRange.Font.Bold = $Bold
    $shape.TextFrame.WordWrap = -1
    $shape.TextFrame.MarginLeft = 4
    $shape.TextFrame.MarginRight = 4
    $shape.TextFrame.MarginTop = 2
    $shape.TextFrame.MarginBottom = 2
    return $shape
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
        [string]$Name
    )
    $rect = $Slide.Shapes.AddShape(5, $Left, $Top, $Width, $Height)
    $rect.Name = $Name
    $rect.Fill.ForeColor.RGB = $C.Panel
    $rect.Fill.Transparency = 0.05
    $rect.Line.ForeColor.RGB = $Accent
    $rect.Line.Weight = 1.25
    $bar = $Slide.Shapes.AddShape(1, $Left, $Top, 5, $Height)
    $bar.Fill.ForeColor.RGB = $Accent
    $bar.Line.Visible = 0
    Add-Text -Slide $Slide -Left ($Left + 13) -Top ($Top + 8) -Width ($Width - 22) -Height 18 -Text $Header -Size ([single]11) -Color $Accent -Bold 1 | Out-Null
    Add-Text -Slide $Slide -Left ($Left + 13) -Top ($Top + 31) -Width ($Width - 24) -Height ($Height - 38) -Text $Body -Size ([single]8.8) -Color $C.White -Bold 0 | Out-Null
}

function Add-GuideSlideBase {
    param([object]$Slide, [int]$SlideNo, [string]$Title, [string]$Subtitle)
    $Slide.FollowMasterBackground = $false
    $Slide.Background.Fill.ForeColor.RGB = $C.Navy
    Add-Text -Slide $Slide -Left 34 -Top 18 -Width 720 -Height 30 -Text $Title -Size ([single]22) -Color $C.White -Bold 1 -Name 'GuideTitle' | Out-Null
    Add-Text -Slide $Slide -Left 36 -Top 53 -Width 790 -Height 22 -Text $Subtitle -Size ([single]9.8) -Color $C.Muted -Bold 0 -Name 'GuideSubtitle' | Out-Null
    $line = $Slide.Shapes.AddShape(1, 36, 82, 858, 2)
    $line.Fill.ForeColor.RGB = $C.Cyan
    $line.Line.Visible = 0
    Add-Text -Slide $Slide -Left 34 -Top 508 -Width 420 -Height 18 -Text 'Future Industrial PLM / Meeting Guide' -Size ([single]7.8) -Color $C.Dim -Bold 0 | Out-Null
    Add-Text -Slide $Slide -Left 872 -Top 508 -Width 38 -Height 18 -Text ([string]$SlideNo) -Size ([single]7.8) -Color $C.Muted -Bold 0 | Out-Null
}

function Set-Notes {
    param([object]$Slide, [string]$Text)
    try {
        $Slide.NotesPage.Shapes.Placeholders(2).TextFrame.TextRange.Text = $Text
    } catch {
        $shape = $Slide.NotesPage.Shapes.AddTextbox(1, 36, 120, 648, 360)
        $shape.TextFrame.TextRange.Text = $Text
    }
}

$pp = New-Object -ComObject PowerPoint.Application
$presentation = $null
try {
    $presentation = $pp.Presentations.Open($v7Deck, $false, $false, $false)
    if ($presentation.Slides.Count -ne 37) {
        throw "Unexpected source slide count: $($presentation.Slides.Count). Expected 37."
    }

    # Mark the source revision without disturbing the cover layout.
    $s1 = $presentation.Slides.Item(1)
    for ($i = 1; $i -le $s1.Shapes.Count; $i++) {
        $shape = $s1.Shapes.Item($i)
        try {
            if (($shape.HasTextFrame -eq -1) -and ($shape.TextFrame.HasText -eq -1)) {
                $txt = $shape.TextFrame.TextRange.Text
                if ($txt -match 'Rev\. V6') {
                    $shape.TextFrame.TextRange.Text = ($txt -replace 'Rev\. V6', 'Rev. V7')
                }
            }
        } catch {}
    }

    # Small navigation cue on the agenda; the main ranges remain unchanged.
    $s2 = $presentation.Slides.Item(2)
    Add-Text -Slide $s2 -Left 540 -Top 72 -Width 360 -Height 16 -Text 'Presenter guide appendix -> p.38-41' -Size ([single]8.5) -Color $C.Cyan -Bold 0 -Name 'PresenterGuidePointer' | Out-Null

    $slide38 = $presentation.Slides.Add($presentation.Slides.Count + 1, 12)
    Add-GuideSlideBase -Slide $slide38 -SlideNo 38 -Title 'Meeting Guide - persuasion storyline' -Subtitle 'Use this as the spoken structure for development and planning teams.'
    Add-Card -Slide $slide38 -Left 40 -Top 112 -Width 405 -Height 110 -Header '1. Start with the gap' -Body 'No single vendor owns the complete loop from engineering change to configuration impact, certified assurance evidence, approval and operations feedback.' -Accent $C.Cyan -Name 'GuideCard1'
    Add-Card -Slide $slide38 -Left 475 -Top 112 -Width 405 -Height 110 -Header '2. Make the strategic move' -Body 'Position AVEVA Marine as the open decision control plane above specialist tools: configuration, impact, evidence, approval and lifecycle learning.' -Accent $C.Teal -Name 'GuideCard2'
    Add-Card -Slide $slide38 -Left 40 -Top 252 -Width 405 -Height 110 -Header '3. Prove feasibility' -Body 'Keep E3D / Marine / Hull authoring on Windows. Host API, SQL, AI Gate, evidence and integration services on Linux. Connect through Local Agent / Plugin adapters.' -Accent $C.Amber -Name 'GuideCard3'
    Add-Card -Slide $slide38 -Left 475 -Top 252 -Width 405 -Height 110 -Header '4. Close with a bounded pilot' -Body 'Ask for Phase 2 configuration core plus one Continuous Naval Assurance pilot slice. The proof is one closed, auditable change-to-evidence loop.' -Accent $C.Cyan -Name 'GuideCard4'
    Add-Text -Slide $slide38 -Left 54 -Top 406 -Width 805 -Height 48 -Text 'Presenter rule: do not sell this as a tool-replacement story. Sell it as a governed, auditable, operations-aware decision loop that makes existing tools more valuable.' -Size ([single]13) -Color $C.White -Bold 1 | Out-Null
    Set-Notes -Slide $slide38 -Text 'Use this slide only if the room needs a clearer storyline. The persuasive arc is gap -> strategic move -> feasible architecture -> bounded pilot decision.'

    $slide39 = $presentation.Slides.Add($presentation.Slides.Count + 1, 12)
    Add-GuideSlideBase -Slide $slide39 -SlideNo 39 -Title 'What development teams must hear' -Subtitle 'The future direction is buildable because the scope is bounded and the interfaces are explicit.'
    Add-Card -Slide $slide39 -Left 38 -Top 108 -Width 270 -Height 276 -Header 'BUILD' -Body "OpenAPI contracts for objects, BOM, effectivity, baseline, change impact and assurance evidence.`r`n`r`nLocal Agent / Plugin callbacks from Windows authoring tools.`r`n`r`nAI Gate provider interface: rules first, RAG next, CONNECT later." -Accent $C.Cyan -Name 'DevBuildCard'
    Add-Card -Slide $slide39 -Left 326 -Top 108 -Width 270 -Height 276 -Header 'VALIDATE' -Body "Seed truth cases for hull, block, compartment, weight, loading case and affected scenarios.`r`n`r`nE2E tests for evaluate -> review -> promote and rollback.`r`n`r`nKPI gates for contract coverage, accuracy, provenance, response time and security." -Accent $C.Teal -Name 'DevValidateCard'
    Add-Card -Slide $slide39 -Left 614 -Top 108 -Width 270 -Height 276 -Header 'AVOID' -Body "No direct Linux hosting assumption for E3D / Marine / Hull.`r`n`r`nNo big-bang replacement of certified solvers.`r`n`r`nNo automatic AI release authority.`r`n`r`nNo unbounded process customization without approval gates." -Accent $C.Red -Name 'DevAvoidCard'
    Add-Text -Slide $slide39 -Left 56 -Top 414 -Width 810 -Height 46 -Text 'Development success means a small number of interfaces become reliable enough that planning, engineering and class reviewers can trust the decision state.' -Size ([single]12.5) -Color $C.White -Bold 1 | Out-Null
    Set-Notes -Slide $slide39 -Text 'For development, stay concrete. The architecture is not a giant PLM rewrite. It is API contracts, schema, adapters, evidence objects, security and tests.'

    $slide40 = $presentation.Slides.Add($presentation.Slides.Count + 1, 12)
    Add-GuideSlideBase -Slide $slide40 -SlideNo 40 -Title 'What planning teams must hear' -Subtitle 'The product story is not only technical; it creates a defensible AVEVA Marine market position.'
    Add-Card -Slide $slide40 -Left 40 -Top 112 -Width 400 -Height 122 -Header 'Product position' -Body 'AVEVA owns the control plane above specialist tools. The customer keeps tool choice, but AVEVA owns the lifecycle decision, evidence and approval layer.' -Accent $C.Cyan -Name 'PlanPositionCard'
    Add-Card -Slide $slide40 -Left 480 -Top 112 -Width 400 -Height 122 -Header 'Customer value' -Body 'Fewer blind handovers, faster change impact review, visible assurance evidence, explainable approvals and a replayable history from design to operations.' -Accent $C.Teal -Name 'PlanValueCard'
    Add-Card -Slide $slide40 -Left 40 -Top 260 -Width 400 -Height 122 -Header 'Roadmap discipline' -Body 'A 26-week gate model prevents wishful scope. Each capability must pass measurable acceptance: contract coverage, accuracy, provenance, response time and release governance.' -Accent $C.Amber -Name 'PlanRoadmapCard'
    Add-Card -Slide $slide40 -Left 480 -Top 260 -Width 400 -Height 122 -Header 'Decision request' -Body 'Approve the bounded Phase 2 + Continuous Naval Assurance increment and agree the owners, KPI thresholds, assumption boundaries and pilot scenario.' -Accent $C.Cyan -Name 'PlanDecisionCard'
    Add-Text -Slide $slide40 -Left 58 -Top 414 -Width 805 -Height 42 -Text 'Planning success means the product direction is easy to explain: AVEVA Marine becomes the trusted system for seeing, approving and learning from every major shipbuilding change.' -Size ([single]12.5) -Color $C.White -Bold 1 | Out-Null
    Set-Notes -Slide $slide40 -Text 'For planning, keep the message tied to market differentiation, customer confidence, scope control and the approval decision requested in the meeting.'

    $slide41 = $presentation.Slides.Add($presentation.Slides.Count + 1, 12)
    Add-GuideSlideBase -Slide $slide41 -SlideNo 41 -Title 'Closing argument - future AVEVA Marine direction' -Subtitle 'The one-sentence story and the final ask.'
    $statement = $slide41.Shapes.AddShape(5, 52, 118, 820, 112)
    $statement.Fill.ForeColor.RGB = $C.Panel2
    $statement.Fill.Transparency = 0.02
    $statement.Line.ForeColor.RGB = $C.Cyan
    $statement.Line.Weight = 1.5
    Add-Text -Slide $slide41 -Left 72 -Top 138 -Width 780 -Height 76 -Text 'AVEVA Marine should become the open, governed control plane that lets shipyards see, explain, approve and replay every engineering state transition across design, assurance and operations.' -Size ([single]18) -Color $C.White -Bold 1 | Out-Null
    Add-Card -Slide $slide41 -Left 52 -Top 262 -Width 255 -Height 130 -Header 'Why it is credible' -Body 'The reference package already proves API-first services, PostgreSQL schema, OIDC security and gate patterns.' -Accent $C.Cyan -Name 'CloseCredible'
    Add-Card -Slide $slide41 -Left 333 -Top 262 -Width 255 -Height 130 -Header 'Why it is differentiated' -Body 'It combines configuration rigor, collaboration context, certified assurance evidence and operations learning.' -Accent $C.Teal -Name 'CloseDifferentiated'
    Add-Card -Slide $slide41 -Left 614 -Top 262 -Width 255 -Height 130 -Header 'What to approve' -Body 'Phase 2 configuration core plus one bounded Continuous Naval Assurance pilot, governed by the WBS and KPI gates.' -Accent $C.Amber -Name 'CloseApprove'
    Add-Text -Slide $slide41 -Left 70 -Top 426 -Width 780 -Height 34 -Text 'Meeting ask: align on product direction, approve pilot scope, assign owners and validate the first customer/class evidence scenario.' -Size ([single]12.5) -Color $C.White -Bold 1 | Out-Null
    Set-Notes -Slide $slide41 -Text 'Use this as the final spoken close if the room needs the whole proposal compressed into one message. The ask is alignment, bounded pilot approval, ownership and first scenario validation.'

    $presentation.Save()
    $presentation.SaveAs($finalPdf, 32)
    $presentation.Close()
    $presentation = $null
} finally {
    if ($presentation -ne $null) { $presentation.Close() }
    if ($pp.Presentations.Count -eq 0) { $pp.Quit() }
}

Copy-Item -LiteralPath $v7Deck -Destination $defaultDeck -Force
Copy-Item -LiteralPath $v7Deck -Destination $finalDeck -Force

$brief = @'
# AVEVA Marine Development Future Direction

## Development and Planning Meeting Brief - English Final Materials

### Meeting Purpose
- Align development and planning teams on the future direction: an open AVEVA Marine PLM control plane above specialist engineering, solver and operations tools.
- Approve a bounded next increment: Phase 2 configuration core plus Continuous Naval Assurance pilot.
- Use objective KPI gates to decide whether each capability is ready to promote.

### Core Thesis
AVEVA Marine should not try to replace every specialist tool. The stronger future position is to own the governed decision loop that connects Windows-based authoring, Linux-based control-plane services, certified solver evidence, human/class approval and operations feedback.

### Persuasion Logic
- For development: the architecture is feasible because it uses bounded FastAPI/OpenAPI contracts, PostgreSQL evidence storage, OIDC/RBAC security, Local Agent/Plugin integration and automated tests.
- For planning: the direction is differentiated because it combines Siemens-grade configuration, Dassault-grade collaboration, NAPA-grade assurance and AVEVA lifecycle context without surrendering the decision layer.
- For risk owners: AI is limited to diagnosis, recommendation and gate-hold support. Certified solvers, engineers and class authorities retain release authority.

### Proposed Future Direction
- Build the Open AVEVA Decision Control Plane: configuration, impact, evidence, approval and operations learning.
- Keep E3D/Marine/Hull authoring on Windows and connect it through Local Agent, Plugin or PML.NET bridge patterns.
- Host API, SQL, AI Gate, evidence and integration services on the Linux control-plane layer.
- Federate NAPA and other certified solvers as sources of calculation evidence, not as systems to replace.
- Start AI with deterministic rule gates and on-prem RAG, then scale toward CONNECT and Unified Engineering AI after governance is validated.

### Meeting Flow
- Slides 3-6 establish current gaps and public product signals.
- Slides 7-15 explain the winning product direction, including AI and lightweight routing.
- Slides 16-27 translate the direction into architecture, workflow and pilot scope.
- Slides 28-34 convert the proposal into WBS, KPI gates, risks and a decision request.
- Slides 35-37 provide glossary support.
- Slides 38-41 provide presenter guidance for persuading development and planning stakeholders.

### Decisions to Secure
- Agree that AVEVA should own the decision/evidence/control plane above specialist tools.
- Approve the Phase 2 configuration core and bounded Continuous Naval Assurance pilot.
- Accept the hybrid boundary: Windows authoring tier plus Linux control plane plus Cloud/CONNECT adapters.
- Confirm that AI remains advisory and gated, never the certified release authority.
- Use the 26-week WBS and KPI model as the delivery and acceptance structure.

### Recommended Closing Message
The next differentiator for AVEVA Marine is not another isolated tool feature. It is a governed, auditable, operations-aware control plane that lets customers see, explain, approve and replay every major engineering state transition. The reference package proves the implementation pattern; the meeting should approve the bounded pilot that proves the product direction.
'@

[System.IO.File]::WriteAllText($briefMd, $brief, [System.Text.Encoding]::UTF8)
[System.IO.File]::WriteAllText($briefTxt, ($brief -replace '#',''), [System.Text.Encoding]::UTF8)

# Export slide PNGs for visual QA.
$pp2 = New-Object -ComObject PowerPoint.Application
$presentation2 = $null
try {
    $presentation2 = $pp2.Presentations.Open($finalDeck, $true, $false, $false)
    for ($i = 1; $i -le $presentation2.Slides.Count; $i++) {
        $pngPath = Join-Path $renderDir ("slide{0:D2}.png" -f $i)
        $presentation2.Slides.Item($i).Export($pngPath, 'PNG', 1600, 900)
    }
    $slideCount = $presentation2.Slides.Count
    $notesCount = 0
    for ($i = 1; $i -le $presentation2.Slides.Count; $i++) {
        try {
            $noteText = $presentation2.Slides.Item($i).NotesPage.Shapes.Placeholders(2).TextFrame.TextRange.Text
            if (($noteText -replace "\s+", '').Length -gt 0) { $notesCount++ }
        } catch {}
    }
    $presentation2.Close()
    $presentation2 = $null
} finally {
    if ($presentation2 -ne $null) { $presentation2.Close() }
    if ($pp2.Presentations.Count -eq 0) { $pp2.Quit() }
}

$pngCount = (Get-ChildItem -LiteralPath $renderDir -Filter '*.png' | Measure-Object).Count
$files = @($v7Deck, $defaultDeck, $finalDeck, $finalPdf, $briefMd, $briefTxt) | ForEach-Object { Get-Item -LiteralPath $_ }

$report = New-Object System.Collections.Generic.List[string]
$report.Add('Future direction meeting materials V7')
$report.Add(("Generated: {0:yyyy-MM-dd HH:mm:ss K}" -f (Get-Date)))
$report.Add(("Slide count: {0}" -f $slideCount))
$report.Add(("Slides with speaker notes: {0}" -f $notesCount))
$report.Add(("PNG render count: {0}" -f $pngCount))
$report.Add(("Render folder: {0}" -f $renderDir))
$report.Add(("Backup folder: {0}" -f $backupDir))
$report.Add('')
$report.Add('Output files:')
foreach ($file in $files) {
    $report.Add(("{0}`t{1}`t{2:yyyy-MM-dd HH:mm:ss}" -f $file.FullName, $file.Length, $file.LastWriteTime))
}
[System.IO.File]::WriteAllLines($reportPath, $report, [System.Text.Encoding]::UTF8)

Get-Content -LiteralPath $reportPath
