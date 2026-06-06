$ErrorActionPreference = 'Stop'

$proposal = 'D:\이력서\AVEVA - Marine Principal Technical Support & Consultant – PLM SME, Busan\Proposal'
$source = Join-Path $proposal 'Future_Industrial_PLM_Meeting_Deck_EN_V7.pptx'
$v8 = Join-Path $proposal 'Future_Industrial_PLM_Meeting_Deck_EN_V8.pptx'
$defaultDeck = Join-Path $proposal 'Future_Industrial_PLM_Meeting_Deck_EN.pptx'
$finalDeck = Join-Path $proposal 'Future_Industrial_PLM_Meeting_Deck_EN_Final.pptx'
$finalPdf = Join-Path $proposal 'Future_Industrial_PLM_Meeting_Deck_EN_Final.pdf'
$workDir = 'C:\Users\namma\.claude\plm_slide_work'
$renderDir = Join-Path $workDir 'future_direction_meeting_v8_reordered_render'
$reportPath = Join-Path $workDir 'future_direction_meeting_v8_reordered_report.txt'
$stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$backupDir = Join-Path $proposal ("_backup_20260606_reorder_v8_" + $stamp)

if (-not (Test-Path -LiteralPath $source)) { throw "Missing source deck: $source" }
if (Test-Path -LiteralPath $renderDir) { Remove-Item -LiteralPath $renderDir -Recurse -Force }
New-Item -ItemType Directory -Path $renderDir | Out-Null
New-Item -ItemType Directory -Path $backupDir | Out-Null

foreach ($p in @($source, $v8, $defaultDeck, $finalDeck, $finalPdf)) {
    if (Test-Path -LiteralPath $p) {
        Copy-Item -LiteralPath $p -Destination (Join-Path $backupDir (Split-Path $p -Leaf)) -Force
    }
}
Copy-Item -LiteralPath $source -Destination $v8 -Force

function Get-ShapeText {
    param([object]$Shape)
    $texts = New-Object System.Collections.Generic.List[string]
    try {
        if ($Shape.Type -eq 6) {
            for ($i = 1; $i -le $Shape.GroupItems.Count; $i++) {
                foreach ($t in (Get-ShapeText -Shape $Shape.GroupItems.Item($i))) { $texts.Add($t) }
            }
            return $texts
        }
    } catch {}
    try {
        if (($Shape.HasTextFrame -eq -1) -and ($Shape.TextFrame.HasText -eq -1)) {
            $txt = ($Shape.TextFrame.TextRange.Text -replace "[`r`n]+", " " -replace "\s+", " ").Trim()
            if ($txt.Length -gt 0) { $texts.Add($txt) }
        }
    } catch {}
    try {
        if ($Shape.HasTable -eq -1) {
            $tbl = $Shape.Table
            for ($r = 1; $r -le $tbl.Rows.Count; $r++) {
                for ($c = 1; $c -le $tbl.Columns.Count; $c++) {
                    try {
                        $txt = ($tbl.Cell($r, $c).Shape.TextFrame.TextRange.Text -replace "[`r`n]+", " " -replace "\s+", " ").Trim()
                        if ($txt.Length -gt 0) { $texts.Add($txt) }
                    } catch {}
                }
            }
        }
    } catch {}
    return $texts
}

function Slide-Contains {
    param([object]$Slide, [string]$Needle)
    for ($i = 1; $i -le $Slide.Shapes.Count; $i++) {
        foreach ($txt in (Get-ShapeText -Shape $Slide.Shapes.Item($i))) {
            if ($txt -like ("*" + $Needle + "*")) { return $true }
        }
    }
    return $false
}

function Find-Slide {
    param([object]$Presentation, [string]$Needle)
    for ($i = 1; $i -le $Presentation.Slides.Count; $i++) {
        $slide = $Presentation.Slides.Item($i)
        if (Slide-Contains -Slide $slide -Needle $Needle) { return $slide }
    }
    throw "Could not find slide containing: $Needle"
}

function Replace-TextInShape {
    param([object]$Shape, [hashtable]$Map)
    try {
        if ($Shape.Type -eq 6) {
            for ($i = 1; $i -le $Shape.GroupItems.Count; $i++) {
                Replace-TextInShape -Shape $Shape.GroupItems.Item($i) -Map $Map
            }
            return
        }
    } catch {}

    try {
        if (($Shape.HasTextFrame -eq -1) -and ($Shape.TextFrame.HasText -eq -1)) {
            $txt = $Shape.TextFrame.TextRange.Text
            foreach ($key in $Map.Keys) {
                if ($txt -eq $key) { $Shape.TextFrame.TextRange.Text = $Map[$key]; return }
                if ($txt -like ("*" + $key + "*")) { $txt = $txt.Replace($key, $Map[$key]) }
            }
            $Shape.TextFrame.TextRange.Text = $txt
        }
    } catch {}
}

function Replace-TextOnSlide {
    param([object]$Slide, [hashtable]$Map)
    for ($i = 1; $i -le $Slide.Shapes.Count; $i++) {
        Replace-TextInShape -Shape $Slide.Shapes.Item($i) -Map $Map
    }
}

function Update-PageNumber {
    param([object]$Slide, [int]$NewNo)
    $candidateTexts = @()
    for ($n = 1; $n -le 60; $n++) { $candidateTexts += [string]$n }
    for ($i = 1; $i -le $Slide.Shapes.Count; $i++) {
        $shape = $Slide.Shapes.Item($i)
        try {
            if (($shape.HasTextFrame -eq -1) -and ($shape.TextFrame.HasText -eq -1)) {
                $txt = ($shape.TextFrame.TextRange.Text -replace "[`r`n]+", " " -replace "\s+", " ").Trim()
                $isLikelyPageNum = ($candidateTexts -contains $txt) -and (($shape.Left -gt 820) -or ($shape.Top -gt 480) -or (($shape.Top -lt 30) -and ($shape.Left -gt 850)))
                if ($isLikelyPageNum) {
                    $shape.TextFrame.TextRange.Text = [string]$NewNo
                }
            }
        } catch {}
    }
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

$desired = @(
    'Company-by-company benchmark',
    'Meeting structure',
    'Meeting Guide - persuasion storyline',
    'What development teams must hear',
    'What planning teams must hear',
    'Current-state comparison',
    'Current PLM landscape',
    'Where we are now',
    'Public product + roadmap signals',
    'Meeting thesis',
    'Strategic gap to fill',
    'What to absorb',
    'Win above NAPA',
    'Win above Siemens',
    'Win above Dassault',
    'AVEVA AI Capabilities',
    'AI deployment path',
    'Lightweight Piping Routing',
    'Target architecture concept',
    'Executable reference package',
    'Hybrid: Windows Authoring + Linux Control Plane + Cloud/CONNECT',
    'Hybrid deployment overview',
    'Phase 2 + Continuous Naval Assurance pilot scope',
    'Requirement',
    'AI-gated flow across AVEVA design tools',
    'Design tool',
    'E3D / Hull design',
    'Continuous Naval Assurance Control Plane',
    'User-Configurable PLM Process',
    'Executable gate flow',
    'Delivery WBS',
    'KPI model',
    'Future-state comparison',
    'Review ownership',
    'Key review risks',
    'Evidence base',
    'Recommended decision',
    'Closing argument',
    'Glossary 1 / 3',
    'Glossary 2 / 3',
    'Glossary 3 / 3'
)

$pp = New-PowerPoint
$presentation = $null
try {
    $presentation = $pp.Presentations.Open($v8, $false, $false, $false)
    if ($presentation.Slides.Count -ne 41) { throw "Expected 41 slides, found $($presentation.Slides.Count)" }

    for ($targetIndex = 1; $targetIndex -le $desired.Count; $targetIndex++) {
        $needle = $desired[$targetIndex - 1]
        $slide = Find-Slide -Presentation $presentation -Needle $needle
        $slide.MoveTo($targetIndex)
    }

    # Refresh title/revision and agenda navigation after the procedural slides move to the front.
    Replace-TextOnSlide -Slide $presentation.Slides.Item(1) -Map @{
        'Rev. V7' = 'Rev. V8'
    }
    Replace-TextOnSlide -Slide $presentation.Slides.Item(2) -Map @{
        'Meeting structure - seven decisions in order' = 'Meeting procedure - development + planning decision flow'
        'Move from facts to ideas, then validate the delivery plan and final decision' = 'Start with the meeting procedure, then move from facts to ideas, delivery plan and final decision'
        'Presenter guide appendix -> p.38-41' = 'Meeting procedure guide -> p.3-5'
        'Abbreviations → Glossary appendix, p.35–37' = 'Abbreviations → Glossary appendix, p.39–41'
        'p.3–5' = 'p.6–8'
        'p.6' = 'p.9'
        'p.7–15' = 'p.10–18'
        'p.16–29' = 'p.19–32'
        'p.30' = 'p.33'
        'p.31–33' = 'p.34–36'
        'p.34' = 'p.37–38'
        'Sequencing rule: WBS and KPI appear only after all proposed capabilities and pilot ideas are consolidated.' = 'Procedure rule: align how the room will decide first; WBS and KPI appear only after capabilities and pilot scope are clear.'
    }

    Replace-TextOnSlide -Slide $presentation.Slides.Item(3) -Map @{
        'Meeting Guide - persuasion storyline' = 'Meeting Procedure - persuasion storyline'
        'Future Industrial PLM / Meeting Guide' = 'Future Industrial PLM / Meeting Procedure'
    }
    Replace-TextOnSlide -Slide $presentation.Slides.Item(4) -Map @{
        'Future Industrial PLM / Meeting Guide' = 'Future Industrial PLM / Meeting Procedure'
    }
    Replace-TextOnSlide -Slide $presentation.Slides.Item(5) -Map @{
        'Future Industrial PLM / Meeting Guide' = 'Future Industrial PLM / Meeting Procedure'
    }
    Replace-TextOnSlide -Slide $presentation.Slides.Item(38) -Map @{
        'Future Industrial PLM / Meeting Guide' = 'Future Industrial PLM / Meeting Procedure'
    }

    for ($i = 1; $i -le $presentation.Slides.Count; $i++) {
        Update-PageNumber -Slide $presentation.Slides.Item($i) -NewNo $i
    }

    # Export and render for QA.
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

Copy-Item -LiteralPath $v8 -Destination $defaultDeck -Force
Copy-Item -LiteralPath $v8 -Destination $finalDeck -Force

# Inspect final order from saved file.
$pp2 = New-PowerPoint
$presentation2 = $null
$orderLines = New-Object System.Collections.Generic.List[string]
$noteCount = 0
try {
    $presentation2 = $pp2.Presentations.Open($finalDeck, $true, $false, $false)
    for ($i = 1; $i -le $presentation2.Slides.Count; $i++) {
        $title = ''
        $slide = $presentation2.Slides.Item($i)
        for ($s = 1; $s -le $slide.Shapes.Count; $s++) {
            foreach ($txt in (Get-ShapeText -Shape $slide.Shapes.Item($s))) {
                if ($txt -ne 'Future Industrial PLM' -and $txt -notmatch '^\d+$') { $title = $txt; break }
            }
            if ($title.Length -gt 0) { break }
        }
        try {
            $nt = $slide.NotesPage.Shapes.Placeholders(2).TextFrame.TextRange.Text
            if (($nt -replace '\s+', '').Length -gt 0) { $noteCount++ }
        } catch {}
        $orderLines.Add(("{0}. {1}" -f $i, $title))
    }
    $slideCount = $presentation2.Slides.Count
    $presentation2.Close()
    $presentation2 = $null
} finally {
    if ($presentation2 -ne $null) { $presentation2.Close() }
    if ($pp2.Presentations.Count -eq 0) { $pp2.Quit() }
}

$pngCount = (Get-ChildItem -LiteralPath $renderDir -Filter '*.png' | Measure-Object).Count
$files = @($v8, $defaultDeck, $finalDeck, $finalPdf) | ForEach-Object { Get-Item -LiteralPath $_ }

$report = New-Object System.Collections.Generic.List[string]
$report.Add('Future Industrial PLM V8 reordered for development/planning meeting procedure')
$report.Add(("Generated: {0:yyyy-MM-dd HH:mm:ss K}" -f (Get-Date)))
$report.Add(("Slide count: {0}" -f $slideCount))
$report.Add(("Slides with notes: {0}" -f $noteCount))
$report.Add(("PNG render count: {0}" -f $pngCount))
$report.Add(("Render folder: {0}" -f $renderDir))
$report.Add(("Backup folder: {0}" -f $backupDir))
$report.Add('')
$report.Add('Saved files:')
foreach ($f in $files) {
    $report.Add(("{0}`t{1}`t{2:yyyy-MM-dd HH:mm:ss}" -f $f.FullName, $f.Length, $f.LastWriteTime))
}
$report.Add('')
$report.Add('Final slide order:')
foreach ($line in $orderLines) { $report.Add($line) }
[System.IO.File]::WriteAllLines($reportPath, $report, [System.Text.Encoding]::UTF8)
Get-Content -LiteralPath $reportPath
