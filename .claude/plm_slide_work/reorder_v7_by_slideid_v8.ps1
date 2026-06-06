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
$backupDir = Join-Path $proposal ("_backup_20260606_reorder_v8_slideid_" + $stamp)

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

function New-Replacement {
    param([string]$From, [string]$To, [string]$Mode)
    return [pscustomobject]@{ From = $From; To = $To; Mode = $Mode }
}

function Replace-OnShape {
    param([object]$Shape, [object[]]$Pairs)
    try {
        if ($Shape.Type -eq 6) {
            for ($i = 1; $i -le $Shape.GroupItems.Count; $i++) {
                Replace-OnShape -Shape $Shape.GroupItems.Item($i) -Pairs $Pairs
            }
            return
        }
    } catch {}
    try {
        if (($Shape.HasTextFrame -eq -1) -and ($Shape.TextFrame.HasText -eq -1)) {
            $txt = $Shape.TextFrame.TextRange.Text
            foreach ($pair in $Pairs) {
                $from = $pair.From
                $to = $pair.To
                $mode = $pair.Mode
                if ($mode -eq 'exact') {
                    if ($txt -eq $from) { $txt = $to }
                } else {
                    if ($txt -like ("*" + $from + "*")) { $txt = $txt.Replace($from, $to) }
                }
            }
            $Shape.TextFrame.TextRange.Text = $txt
        }
    } catch {}
}

function Replace-OnSlide {
    param([object]$Slide, [object[]]$Pairs)
    for ($i = 1; $i -le $Slide.Shapes.Count; $i++) {
        Replace-OnShape -Shape $Slide.Shapes.Item($i) -Pairs $Pairs
    }
}

function Update-PageNumber {
    param([object]$Slide, [int]$NewNo)
    for ($i = 1; $i -le $Slide.Shapes.Count; $i++) {
        $shape = $Slide.Shapes.Item($i)
        try {
            if (($shape.HasTextFrame -eq -1) -and ($shape.TextFrame.HasText -eq -1)) {
                $txt = ($shape.TextFrame.TextRange.Text -replace "[`r`n]+", " " -replace "\s+", " ").Trim()
                $numeric = 0
                $isNumber = [int]::TryParse($txt, [ref]$numeric)
                $isLikelyPageNum = $isNumber -and ($numeric -ge 1) -and ($numeric -le 60) -and (($shape.Left -gt 820) -or ($shape.Top -gt 480) -or (($shape.Top -lt 30) -and ($shape.Left -gt 850)))
                if ($isLikelyPageNum) { $shape.TextFrame.TextRange.Text = [string]$NewNo }
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

# V7 source order:
# 1 cover, 2 agenda, 3-37 main/glossary, 38 guide, 39 dev, 40 planning, 41 closing argument.
# V8 meeting order: procedure first, closing before glossary.
$newOrderFromOldIndex = @(
    1, 2, 38, 39, 40,
    3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15,
    16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29,
    30, 31, 32, 33, 34, 41, 35, 36, 37
)

$pp = New-PowerPoint
$presentation = $null
try {
    $presentation = $pp.Presentations.Open($v8, $false, $false, $false)
    if ($presentation.Slides.Count -ne 41) { throw "Expected 41 slides, found $($presentation.Slides.Count)" }

    $idByOldIndex = @{}
    for ($i = 1; $i -le $presentation.Slides.Count; $i++) {
        $idByOldIndex[$i] = $presentation.Slides.Item($i).SlideID
    }

    for ($targetIndex = 1; $targetIndex -le $newOrderFromOldIndex.Count; $targetIndex++) {
        $oldIndex = $newOrderFromOldIndex[$targetIndex - 1]
        $slide = $presentation.Slides.FindBySlideID($idByOldIndex[$oldIndex])
        $slide.MoveTo($targetIndex)
    }

    Replace-OnSlide -Slide $presentation.Slides.Item(1) -Pairs @(
        (New-Replacement -From 'Rev. V7' -To 'Rev. V8' -Mode 'contains')
    )

    Replace-OnSlide -Slide $presentation.Slides.Item(2) -Pairs @(
        (New-Replacement -From 'Meeting structure - seven decisions in order' -To 'Meeting procedure - development + planning decision flow' -Mode 'exact'),
        (New-Replacement -From 'Move from facts to ideas, then validate the delivery plan and final decision' -To 'Start with the meeting procedure, then move from facts to ideas, delivery plan and final decision' -Mode 'exact'),
        (New-Replacement -From 'Presenter guide appendix -> p.38-41' -To 'Meeting procedure guide -> p.3-5' -Mode 'exact'),
        (New-Replacement -From 'Abbreviations → Glossary appendix, p.35–37' -To 'Abbreviations → Glossary appendix, p.39–41' -Mode 'exact'),
        (New-Replacement -From 'p.3–5' -To 'p.6–8' -Mode 'exact'),
        (New-Replacement -From 'p.6' -To 'p.9' -Mode 'exact'),
        (New-Replacement -From 'p.7–15' -To 'p.10–18' -Mode 'exact'),
        (New-Replacement -From 'p.16–29' -To 'p.19–32' -Mode 'exact'),
        (New-Replacement -From 'p.30' -To 'p.33' -Mode 'exact'),
        (New-Replacement -From 'p.31–33' -To 'p.34–36' -Mode 'exact'),
        (New-Replacement -From 'p.34' -To 'p.37–38' -Mode 'exact'),
        (New-Replacement -From 'Sequencing rule: WBS and KPI appear only after all proposed capabilities and pilot ideas are consolidated.' -To 'Procedure rule: align how the room will decide first; WBS and KPI appear only after capabilities and pilot scope are clear.' -Mode 'exact')
    )

    Replace-OnSlide -Slide $presentation.Slides.Item(3) -Pairs @(
        (New-Replacement -From 'Meeting Guide - persuasion storyline' -To 'Meeting Procedure - persuasion storyline' -Mode 'exact'),
        (New-Replacement -From 'Future Industrial PLM / Meeting Guide' -To 'Future Industrial PLM / Meeting Procedure' -Mode 'exact')
    )
    Replace-OnSlide -Slide $presentation.Slides.Item(4) -Pairs @((New-Replacement -From 'Future Industrial PLM / Meeting Guide' -To 'Future Industrial PLM / Meeting Procedure' -Mode 'exact'))
    Replace-OnSlide -Slide $presentation.Slides.Item(5) -Pairs @((New-Replacement -From 'Future Industrial PLM / Meeting Guide' -To 'Future Industrial PLM / Meeting Procedure' -Mode 'exact'))
    Replace-OnSlide -Slide $presentation.Slides.Item(38) -Pairs @((New-Replacement -From 'Future Industrial PLM / Meeting Guide' -To 'Future Industrial PLM / Meeting Procedure' -Mode 'exact'))

    for ($i = 1; $i -le $presentation.Slides.Count; $i++) {
        Update-PageNumber -Slide $presentation.Slides.Item($i) -NewNo $i
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

Copy-Item -LiteralPath $v8 -Destination $defaultDeck -Force
Copy-Item -LiteralPath $v8 -Destination $finalDeck -Force

$pp2 = New-PowerPoint
$presentation2 = $null
$orderLines = New-Object System.Collections.Generic.List[string]
$noteCount = 0
try {
    $presentation2 = $pp2.Presentations.Open($finalDeck, $true, $false, $false)
    for ($i = 1; $i -le $presentation2.Slides.Count; $i++) {
        $slide = $presentation2.Slides.Item($i)
        $title = ''
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
$report.Add('Future Industrial PLM V8 reordered by SlideID for development/planning meeting procedure')
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
