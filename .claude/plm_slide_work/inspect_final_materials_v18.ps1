$ErrorActionPreference = 'Stop'

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

function Get-WordApplication {
    try {
        $word = New-Object -ComObject Word.Application
        $word.Visible = $false
        return $word
    } catch {
        $exe = 'C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE'
        if (-not (Test-Path -LiteralPath $exe)) { throw }
        Start-Process -FilePath $exe -WindowStyle Hidden
        Start-Sleep -Seconds 8
        $word = [Runtime.InteropServices.Marshal]::GetActiveObject('Word.Application')
        $word.Visible = $false
        return $word
    }
}

function Normalize-Text($text) {
    return (($text -replace '[\r\n\a\x07]+', ' ') -replace '\s+', ' ').Trim()
}

$proposal = 'D:\이력서\AVEVA - Marine Principal Technical Support & Consultant – PLM SME, Busan\Proposal'
$deckPath = Join-Path $proposal 'Future_Industrial_PLM_Meeting_Deck_EN_V18.pptx'
$guideEnDocx = Join-Path $proposal 'Future_Industrial_PLM_Presenter_Facilitation_Guide_EN.docx'
$guideEnPdf = Join-Path $proposal 'Future_Industrial_PLM_Presenter_Facilitation_Guide_EN.pdf'
$guideKoDocx = Join-Path $proposal 'Future_Industrial_PLM_Presenter_Facilitation_Guide_KO.docx'
$guideKoPdf = Join-Path $proposal 'Future_Industrial_PLM_Presenter_Facilitation_Guide_KO.pdf'

$work = Join-Path $env:USERPROFILE '.claude\plm_slide_work'
$reportPath = Join-Path $work 'inspect_final_materials_v18_report.txt'
$deckDumpPath = Join-Path $work 'inspect_final_materials_v18_deck_text.txt'
$guideEnTextPath = Join-Path $work 'inspect_final_materials_v18_guide_en_text.txt'
$guideKoTextPath = Join-Path $work 'inspect_final_materials_v18_guide_ko_text.txt'

$lines = New-Object System.Collections.Generic.List[string]
$deckDump = New-Object System.Collections.Generic.List[string]

$targets = @($deckPath, $guideEnDocx, $guideEnPdf, $guideKoDocx, $guideKoPdf)
$lines.Add('FILE METADATA')
foreach ($path in $targets) {
    if (Test-Path -LiteralPath $path) {
        $item = Get-Item -LiteralPath $path
        $lines.Add(('{0} | {1} bytes | {2:yyyy-MM-dd HH:mm:ss}' -f $item.FullName, $item.Length, $item.LastWriteTime))
    } else {
        $lines.Add(('MISSING: {0}' -f $path))
    }
}

$ppt = $null
$pres = $null
try {
    $ppt = Get-PowerPointApplication
    try { $ppt.Visible = [Microsoft.Office.Core.MsoTriState]::msoFalse } catch {}
    $pres = $ppt.Presentations.Open($deckPath, [Microsoft.Office.Core.MsoTriState]::msoTrue, [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoFalse)

    $lines.Add('')
    $lines.Add('DECK INSPECTION')
    $lines.Add(('Slides: {0}' -f $pres.Slides.Count))
    $lines.Add(('Slide size: {0:n1} x {1:n1}' -f $pres.PageSetup.SlideWidth, $pres.PageSetup.SlideHeight))

    $slidesWithoutNotes = New-Object System.Collections.Generic.List[int]
    $slideTitles = New-Object System.Collections.Generic.List[string]
    $revisionText = ''
    $allDeckText = New-Object System.Collections.Generic.List[string]

    for ($s = 1; $s -le $pres.Slides.Count; $s++) {
        $slide = $pres.Slides.Item($s)
        $texts = New-Object System.Collections.Generic.List[string]
        $title = ''
        for ($i = 1; $i -le $slide.Shapes.Count; $i++) {
            $shape = $slide.Shapes.Item($i)
            try {
                if ($shape.HasTextFrame -and $shape.TextFrame.HasText) {
                    $t = Normalize-Text $shape.TextFrame.TextRange.Text
                    if ($t.Length -gt 0) {
                        $texts.Add($t)
                        $allDeckText.Add($t)
                        if ($s -eq 1 -and $t -match 'Rev\. V\d+') { $revisionText = $t }
                        if ($title.Length -eq 0 -and $shape.Top -lt 90 -and $t.Length -gt 1 -and $t -notmatch '^\d+$' -and $t -notmatch '^Future Industrial PLM$') {
                            $title = $t
                        }
                    }
                }
            } catch {}
        }
        if ($title.Length -eq 0 -and $texts.Count -gt 0) { $title = $texts[0] }

        $noteText = ''
        try {
            for ($n = 1; $n -le $slide.NotesPage.Shapes.Count; $n++) {
                $ns = $slide.NotesPage.Shapes.Item($n)
                if ($ns.HasTextFrame -and $ns.TextFrame.HasText) {
                    $noteText += ' ' + (Normalize-Text $ns.TextFrame.TextRange.Text)
                }
            }
        } catch {}
        if ([string]::IsNullOrWhiteSpace($noteText)) {
            $slidesWithoutNotes.Add($s)
        }

        $slideTitles.Add(('p.{0}: {1}' -f $s, $title))
        $deckDump.Add(('--- SLIDE {0}: {1}' -f $s, $title))
        foreach ($t in $texts) { $deckDump.Add($t) }
        $noteOut = $noteText
        if ([string]::IsNullOrWhiteSpace($noteOut)) { $noteOut = '<EMPTY>' }
        $deckDump.Add(('NOTES: {0}' -f $noteOut))
    }

    $deckBlob = ($allDeckText -join ' ')
    $lines.Add(('Cover revision text: {0}' -f $revisionText))
    $slidesWithoutNotesText = 'none'
    if ($slidesWithoutNotes.Count -ne 0) { $slidesWithoutNotesText = ($slidesWithoutNotes -join ', ') }
    $lines.Add(('Slides without notes: {0}' -f $slidesWithoutNotesText))
    $lines.Add(('Contains V17 text refs: {0}' -f (($deckBlob -match 'V17').ToString())))
    $lines.Add(('Contains V18 text refs: {0}' -f (($deckBlob -match 'V18').ToString())))
    $lines.Add(('Contains p.40-42 glossary refs: {0}' -f (($deckBlob -match 'p\.40.?42|p\.40–42|p\.40-42').ToString())))
    $lines.Add(('Contains p.39-41 glossary refs: {0}' -f (($deckBlob -match 'p\.39.?41|p\.39–41|p\.39-41').ToString())))
    $lines.Add('Slide titles:')
    foreach ($entry in $slideTitles) { $lines.Add(('  {0}' -f $entry)) }
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

$word = $null
try {
    $word = Get-WordApplication
    foreach ($docInfo in @(
        @{ Label='EN GUIDE'; Path=$guideEnDocx; TextPath=$guideEnTextPath },
        @{ Label='KO GUIDE'; Path=$guideKoDocx; TextPath=$guideKoTextPath }
    )) {
        $doc = $null
        try {
            $doc = $word.Documents.Open($docInfo.Path, $false, $true)
            $txt = $doc.Content.Text
            $norm = Normalize-Text $txt
            $txt | Set-Content -LiteralPath $docInfo.TextPath -Encoding UTF8

            $lines.Add('')
            $lines.Add($docInfo.Label)
            $lines.Add(('Path: {0}' -f $docInfo.Path))
            $lines.Add(('Pages: {0}' -f $doc.ComputeStatistics(2)))
            $lines.Add(('Paragraphs: {0}' -f $doc.Paragraphs.Count))
            $lines.Add(('Tables: {0}' -f $doc.Tables.Count))
            $lines.Add(('Characters: {0}' -f $txt.Length))
            $lines.Add(('Markdown artifacts (** / #### / [[PB]]): {0}' -f (($norm -match '\*\*|####|\[\[PB\]\]').ToString())))
            $lines.Add(('Mentions V17: {0}' -f (($norm -match 'V17').ToString())))
            $lines.Add(('Mentions V18: {0}' -f (($norm -match 'V18').ToString())))
            $lines.Add(('Mentions 42 slides/pages: {0}' -f (($norm -match '42').ToString())))
            $lines.Add(('Mentions glossary p.40-42: {0}' -f (($norm -match 'p\.40.?42|p\.40–42|p\.40-42').ToString())))
            $lines.Add(('Mentions deck file V18: {0}' -f (($norm -match 'Future_Industrial_PLM_Meeting_Deck_EN_V18').ToString())))

            $headings = New-Object System.Collections.Generic.List[string]
            for ($p = 1; $p -le $doc.Paragraphs.Count; $p++) {
                $pt = Normalize-Text $doc.Paragraphs.Item($p).Range.Text
                if ($pt.Length -gt 0 -and ($pt -match '^[0-9]+\.' -or $pt -match '^(How to use|Meeting snapshot|Audience map|Delivery principles|Opening script|Section-by-section|Anticipated|Closing script|Facilitation logistics|Quick reference|사용 방법|회의 스냅샷|청중 맵|전달 원칙|오프닝 스크립트|슬라이드별|예상 질문|마무리 스크립트|진행 물류|퀵 레퍼런스)')) {
                    $headings.Add($pt)
                }
            }
            $lines.Add('Detected headings / section markers:')
            foreach ($h in ($headings | Select-Object -First 30)) { $lines.Add(('  {0}' -f $h)) }
        } finally {
            if ($doc -ne $null) {
                $doc.Close([ref]$false)
                [void][Runtime.InteropServices.Marshal]::ReleaseComObject($doc)
            }
        }
    }
} finally {
    if ($word -ne $null) {
        if ($word.Documents.Count -eq 0) { $word.Quit() }
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($word)
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

$lines | Set-Content -LiteralPath $reportPath -Encoding UTF8
$deckDump | Set-Content -LiteralPath $deckDumpPath -Encoding UTF8
$lines | ForEach-Object { Write-Output $_ }
