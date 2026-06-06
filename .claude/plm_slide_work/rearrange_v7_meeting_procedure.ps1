$ErrorActionPreference = 'Stop'

$proposal = 'D:\이력서\AVEVA - Marine Principal Technical Support & Consultant – PLM SME, Busan\Proposal'
$v7Deck = Join-Path $proposal 'Future_Industrial_PLM_Meeting_Deck_EN_V7.pptx'
$v8Deck = Join-Path $proposal 'Future_Industrial_PLM_Meeting_Deck_EN_V8.pptx'
$defaultDeck = Join-Path $proposal 'Future_Industrial_PLM_Meeting_Deck_EN.pptx'
$finalDeck = Join-Path $proposal 'Future_Industrial_PLM_Meeting_Deck_EN_Final.pptx'
$finalPdf = Join-Path $proposal 'Future_Industrial_PLM_Meeting_Deck_EN_Final.pdf'
$workDir = 'C:\Users\namma\.claude\plm_slide_work'
$renderDir = Join-Path $workDir 'v8_meeting_procedure_render'
$reportPath = Join-Path $workDir 'v8_meeting_procedure_report.txt'
$stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$backupDir = Join-Path $proposal ("_backup_20260606_v7_rearrange_meeting_procedure_" + $stamp)

if (-not (Test-Path -LiteralPath $v7Deck)) { throw "V7 deck not found: $v7Deck" }
if (-not (Test-Path -LiteralPath $workDir)) { New-Item -ItemType Directory -Path $workDir | Out-Null }
if (Test-Path -LiteralPath $renderDir) {
    $resolved = (Resolve-Path -LiteralPath $renderDir).Path
    if (-not $resolved.StartsWith($workDir, [System.StringComparison]::OrdinalIgnoreCase)) {
        throw "Refusing to clear unexpected render path: $resolved"
    }
    Remove-Item -LiteralPath $renderDir -Recurse -Force
}
New-Item -ItemType Directory -Path $renderDir | Out-Null
New-Item -ItemType Directory -Path $backupDir | Out-Null

foreach ($path in @($v7Deck, $v8Deck, $defaultDeck, $finalDeck, $finalPdf)) {
    if (Test-Path -LiteralPath $path) {
        Copy-Item -LiteralPath $path -Destination (Join-Path $backupDir (Split-Path $path -Leaf)) -Force
    }
}

function Get-SlideText {
    param([object]$Slide)
    $texts = New-Object System.Collections.Generic.List[string]
    for ($i = 1; $i -le $Slide.Shapes.Count; $i++) {
        $shape = $Slide.Shapes.Item($i)
        try {
            if (($shape.HasTextFrame -eq -1) -and ($shape.TextFrame.HasText -eq -1)) {
                $txt = ($shape.TextFrame.TextRange.Text -replace "[`r`n]+", " " -replace "\s+", " ").Trim()
                if ($txt.Length -gt 0) { $texts.Add($txt) }
            }
        } catch {}
    }
    return ($texts -join ' | ')
}

function Find-SlideByTitle {
    param([object]$Presentation, [string]$Title)
    for ($i = 1; $i -le $Presentation.Slides.Count; $i++) {
        $txt = Get-SlideText -Slide $Presentation.Slides.Item($i)
        if ($txt -like ("*" + $Title + "*")) {
            return $Presentation.Slides.Item($i)
        }
    }
    throw "Slide not found: $Title"
}

function Set-ShapeTextIfMatch {
    param([object]$Shape, [hashtable]$Map)
    try {
        if ($Shape.Type -eq 6) {
            for ($i = 1; $i -le $Shape.GroupItems.Count; $i++) {
                Set-ShapeTextIfMatch -Shape $Shape.GroupItems.Item($i) -Map $Map
            }
            return
        }
    } catch {}

    try {
        if (($Shape.HasTextFrame -eq -1) -and ($Shape.TextFrame.HasText -eq -1)) {
            $txt = $Shape.TextFrame.TextRange.Text
            $norm = ($txt -replace "[`r`n]+", " " -replace "\s+", " ").Trim()
            if ($Map.ContainsKey($norm)) {
                $Shape.TextFrame.TextRange.Text = $Map[$norm]
            } elseif ($norm -like 'Abbreviations*Glossary appendix*') {
                $Shape.TextFrame.TextRange.Text = 'Abbreviations -> Glossary appendix, p.39-41'
            } elseif ($norm -like 'Presenter guide appendix*') {
                $Shape.TextFrame.TextRange.Text = 'Meeting guide: p.3-5 | Closing argument: p.38'
            }
        }
    } catch {}
}

function Update-PageNumberShape {
    param([object]$Shape, [int]$SlideNo)
    try {
        if ($Shape.Type -eq 6) {
            for ($i = 1; $i -le $Shape.GroupItems.Count; $i++) {
                Update-PageNumberShape -Shape $Shape.GroupItems.Item($i) -SlideNo $SlideNo
            }
            return
        }
    } catch {}

    try {
        if (($Shape.HasTextFrame -eq -1) -and ($Shape.TextFrame.HasText -eq -1)) {
            $txt = ($Shape.TextFrame.TextRange.Text -replace "[`r`n]+", " " -replace "\s+", " ").Trim()
            $fontSize = [double]$Shape.TextFrame.TextRange.Font.Size
            $nearPageZone = (($Shape.Left -gt 840) -or ($Shape.Top -lt 25) -or (($Shape.Left -gt 800) -and ($Shape.Top -gt 490)))
            if (($txt -match '^\d+$') -and ($fontSize -le 10.5) -and $nearPageZone) {
                $Shape.TextFrame.TextRange.Text = [string]$SlideNo
            }
        }
    } catch {}
}

function Set-SlideNotes {
    param([object]$Slide, [string]$Text)
    try {
        $Slide.NotesPage.Shapes.Placeholders(2).TextFrame.TextRange.Text = $Text
    } catch {
        $shape = $Slide.NotesPage.Shapes.AddTextbox(1, 36, 120, 648, 360)
        $shape.TextFrame.TextRange.Text = $Text
    }
}

function New-PowerPointApplication {
    try {
        return New-Object -ComObject PowerPoint.Application
    } catch {
        $powerpnt = 'C:\Program Files\Microsoft Office\root\Office16\POWERPNT.EXE'
        if (Test-Path -LiteralPath $powerpnt) {
            Start-Process -FilePath $powerpnt -ArgumentList '/automation' -WindowStyle Hidden
            Start-Sleep -Seconds 4
        }
        return New-Object -ComObject PowerPoint.Application
    }
}

$pp = New-PowerPointApplication
$presentation = $null
try {
    $presentation = $pp.Presentations.Open($v7Deck, $false, $false, $false)
    if ($presentation.Slides.Count -ne 41) {
        throw "Unexpected V7 slide count: $($presentation.Slides.Count). Expected 41."
    }

    (Find-SlideByTitle -Presentation $presentation -Title 'Meeting Guide - persuasion storyline').MoveTo(3)
    (Find-SlideByTitle -Presentation $presentation -Title 'What development teams must hear').MoveTo(4)
    (Find-SlideByTitle -Presentation $presentation -Title 'What planning teams must hear').MoveTo(5)
    (Find-SlideByTitle -Presentation $presentation -Title 'Closing argument - future AVEVA Marine direction').MoveTo(38)

    # Cover revision marker.
    for ($i = 1; $i -le $presentation.Slides.Item(1).Shapes.Count; $i++) {
        $shape = $presentation.Slides.Item(1).Shapes.Item($i)
        try {
            if (($shape.HasTextFrame -eq -1) -and ($shape.TextFrame.HasText -eq -1)) {
                $txt = $shape.TextFrame.TextRange.Text
                if ($txt -match 'Rev\. V7') {
                    $shape.TextFrame.TextRange.Text = ($txt -replace 'Rev\. V7', 'Rev. V8')
                }
            }
        } catch {}
    }

    $agendaMap = @{
        'Meeting structure - seven decisions in order' = 'Meeting procedure - development + planning alignment'
        'Move from facts to ideas, then validate the delivery plan and final decision' = 'Align the room first, then move from facts to future direction, delivery procedure, review and decision.'
        'CURRENT STATE' = 'MEETING PROCEDURE'
        "Competitor vs AVEVA comparison, PLM position and today's executable baseline" = 'Persuasion storyline and development/planning lens before details'
        'PUBLIC SIGNALS' = 'CURRENT STATE + SIGNALS'
        'Official product and roadmap signals; confirmed versus directional' = 'Competitor gap, PLM landscape, Phase 1 baseline and public signals'
        'WINNING IDEAS' = 'STRATEGIC DIRECTION'
        'Meeting thesis + win-above NAPA / Siemens / Dassault patterns' = 'Win-above patterns plus AI and lightweight routing direction'
        'PROPOSAL + DELIVERY' = 'DELIVERY DESIGN'
        'Architecture, workflows and pilot scope - then WBS and KPI gates' = 'Architecture, executable package, hybrid model and pilot workflows'
        'FUTURE-STATE COMPARE' = 'PLAN + KPI'
        'Re-score AVEVA after the proposed ideas are applied' = 'WBS timeline and objective acceptance model'
        'REVIEW' = 'REVIEW + RISK'
        'Ownership, risks, evidence and assumptions requiring agreement' = 'Future-state score, ownership, risk and evidence boundaries'
        'CLOSE' = 'DECISION + CLOSE'
        'Approve the next increment and the assurance pilot' = 'Approval ask and final future-direction argument'
        'Sequencing rule: WBS and KPI appear only after all proposed capabilities and pilot ideas are consolidated.' = 'Procedure rule: align the room first, then move from facts -> direction -> implementation -> governance -> decision.'
        'p.3–5' = 'p.3-5'
        'p.6' = 'p.6-9'
        'p.7–15' = 'p.10-18'
        'p.16–29' = 'p.19-30'
        'p.30' = 'p.31-32'
        'p.31–33' = 'p.33-36'
        'p.34' = 'p.37-38'
    }

    $slide2 = $presentation.Slides.Item(2)
    for ($i = 1; $i -le $slide2.Shapes.Count; $i++) {
        Set-ShapeTextIfMatch -Shape $slide2.Shapes.Item($i) -Map $agendaMap
    }
    Set-SlideNotes -Slide $slide2 -Text 'Use this agenda as the meeting procedure. Start by aligning development and planning teams on how to evaluate the proposal, then move through current-state facts, strategic direction, implementation procedure, KPI governance and the final decision.'

    # Update all visible page number shapes after the slide move.
    for ($s = 1; $s -le $presentation.Slides.Count; $s++) {
        $slide = $presentation.Slides.Item($s)
        for ($i = 1; $i -le $slide.Shapes.Count; $i++) {
            Update-PageNumberShape -Shape $slide.Shapes.Item($i) -SlideNo $s
        }
    }

    $presentation.Save()
    $presentation.SaveAs($v8Deck, 24)
    $presentation.SaveAs($finalPdf, 32)
    $presentation.Close()
    $presentation = $null
} finally {
    if ($presentation -ne $null) { $presentation.Close() }
    if ($pp.Presentations.Count -eq 0) { $pp.Quit() }
}

Copy-Item -LiteralPath $v7Deck -Destination $defaultDeck -Force
Copy-Item -LiteralPath $v7Deck -Destination $finalDeck -Force

$pp2 = New-PowerPointApplication
$presentation2 = $null
try {
    $presentation2 = $pp2.Presentations.Open($finalDeck, $true, $false, $false)
    $orderLines = New-Object System.Collections.Generic.List[string]
    $notesCount = 0
    for ($s = 1; $s -le $presentation2.Slides.Count; $s++) {
        $slide = $presentation2.Slides.Item($s)
        $png = Join-Path $renderDir ("slide{0:D2}.png" -f $s)
        $slide.Export($png, 'PNG', 1600, 900)
        $txt = Get-SlideText -Slide $slide
        $title = (($txt -split '\|') | Select-Object -First 1).Trim()
        $orderLines.Add(("{0}. {1}" -f $s, $title))
        try {
            $noteText = $slide.NotesPage.Shapes.Placeholders(2).TextFrame.TextRange.Text
            if (($noteText -replace "\s+", '').Length -gt 0) { $notesCount++ }
        } catch {}
    }
    $slideCount = $presentation2.Slides.Count
    $presentation2.Close()
    $presentation2 = $null
} finally {
    if ($presentation2 -ne $null) { $presentation2.Close() }
    if ($pp2.Presentations.Count -eq 0) { $pp2.Quit() }
}

$pngCount = (Get-ChildItem -LiteralPath $renderDir -Filter '*.png' | Measure-Object).Count
$files = @($v7Deck, $v8Deck, $defaultDeck, $finalDeck, $finalPdf) | ForEach-Object { Get-Item -LiteralPath $_ }

$report = New-Object System.Collections.Generic.List[string]
$report.Add('V8 meeting procedure rearrangement')
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
$report.Add('')
$report.Add('New slide order:')
foreach ($line in $orderLines) { $report.Add($line) }
[System.IO.File]::WriteAllLines($reportPath, $report, [System.Text.Encoding]::UTF8)

Get-Content -LiteralPath $reportPath
