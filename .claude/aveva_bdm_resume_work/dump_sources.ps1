$ErrorActionPreference = "Stop"

function Normalize-Text {
    param([string]$Text)
    if ($null -eq $Text) { return "" }
    return (($Text -replace "[`r`n`a\x07]", " ") -replace "\s+", " ").Trim()
}

function Ensure-Parent {
    param([string]$Path)
    $dir = Split-Path -Parent $Path
    if (-not (Test-Path -LiteralPath $dir)) {
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
    }
}

function Dump-WordDocument {
    param(
        [string]$Path,
        [string]$OutPath
    )

    Ensure-Parent -Path $OutPath
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $doc = $null
    try {
        $doc = $word.Documents.Open($Path, $false, $true)
        $lines = New-Object System.Collections.Generic.List[string]
        $lines.Add("PATH: $Path")
        $lines.Add("PAGES: " + $doc.ComputeStatistics(2))
        $lines.Add("WORDS: " + $doc.ComputeStatistics(0))
        $sec = $doc.Sections.Item(1)
        $ps = $sec.PageSetup
        $lines.Add(("PAGE_SETUP: W={0}; H={1}; Top={2}; Bottom={3}; Left={4}; Right={5}" -f $ps.PageWidth, $ps.PageHeight, $ps.TopMargin, $ps.BottomMargin, $ps.LeftMargin, $ps.RightMargin))
        $lines.Add("TABLES: " + $doc.Tables.Count)
        $lines.Add("")
        $lines.Add("== PARAGRAPHS ==")
        $i = 0
        foreach ($p in $doc.Paragraphs) {
            $i++
            $txt = Normalize-Text $p.Range.Text
            if ($txt.Length -eq 0) { continue }
            $styleName = ""
            try { $styleName = $p.Range.Style.NameLocal } catch {}
            $fontName = ""
            try { $fontName = $p.Range.Font.Name } catch {}
            $fontSize = ""
            try { $fontSize = [string]$p.Range.Font.Size } catch {}
            $bold = ""
            try { $bold = [string]$p.Range.Font.Bold } catch {}
            $lines.Add(("{0:0000} | Style={1} | Font={2} | Size={3} | Bold={4} | {5}" -f $i, $styleName, $fontName, $fontSize, $bold, $txt))
        }
        $lines.Add("")
        $lines.Add("== TABLE CELLS ==")
        $t = 0
        foreach ($tb in $doc.Tables) {
            $t++
            $lines.Add("-- TABLE $t --")
            $c = 0
            foreach ($cell in $tb.Range.Cells) {
                $c++
                $txt = Normalize-Text $cell.Range.Text
                if ($txt.Length -gt 0) {
                    $lines.Add(("T{0:00} C{1:000}: {2}" -f $t, $c, $txt))
                }
            }
        }
        $lines.Add("")
        $lines.Add("== FULL TEXT ==")
        $lines.Add($doc.Content.Text)
        [System.IO.File]::WriteAllLines($OutPath, $lines, [System.Text.Encoding]::UTF8)
    } finally {
        if ($doc -ne $null) {
            try { $doc.Close([ref]$false) } catch {}
        }
        if ($word -ne $null -and $word.Documents.Count -eq 0) {
            try { $word.Quit() } catch {}
        }
    }
}

function Dump-PowerPoint {
    param(
        [string]$Path,
        [string]$OutPath
    )

    Ensure-Parent -Path $OutPath
    $pp = New-Object -ComObject PowerPoint.Application
    $pres = $null
    try {
        $pres = $pp.Presentations.Open($Path, $true, $false, $false)
        $lines = New-Object System.Collections.Generic.List[string]
        $lines.Add("PATH: $Path")
        $lines.Add("SLIDES: " + $pres.Slides.Count)
        $lines.Add(("SLIDE_SIZE: W={0}; H={1}" -f $pres.PageSetup.SlideWidth, $pres.PageSetup.SlideHeight))
        foreach ($slide in $pres.Slides) {
            $lines.Add("")
            $lines.Add(("== SLIDE {0}: {1} ==" -f $slide.SlideIndex, (Normalize-Text $slide.Name)))
            $shapeIndex = 0
            foreach ($shape in $slide.Shapes) {
                $shapeIndex++
                try {
                    if ($shape.HasTextFrame -and $shape.TextFrame.HasText) {
                        $txt = Normalize-Text $shape.TextFrame.TextRange.Text
                        if ($txt.Length -gt 0) {
                            $lines.Add(("S{0:00} Shape{1:000} {2}: {3}" -f $slide.SlideIndex, $shapeIndex, $shape.Name, $txt))
                        }
                    }
                } catch {}
                try {
                    if ($shape.HasTable) {
                        $rows = $shape.Table.Rows.Count
                        $cols = $shape.Table.Columns.Count
                        for ($r = 1; $r -le $rows; $r++) {
                            $cells = New-Object System.Collections.Generic.List[string]
                            for ($c = 1; $c -le $cols; $c++) {
                                $cellText = Normalize-Text $shape.Table.Cell($r,$c).Shape.TextFrame.TextRange.Text
                                $cells.Add($cellText)
                            }
                            $lines.Add(("S{0:00} Table {1} R{2}: {3}" -f $slide.SlideIndex, $shape.Name, $r, ([string]::Join(" | ", $cells))))
                        }
                    }
                } catch {}
            }
            try {
                if ($slide.NotesPage.Shapes.Count -gt 0) {
                    foreach ($noteShape in $slide.NotesPage.Shapes) {
                        if ($noteShape.HasTextFrame -and $noteShape.TextFrame.HasText) {
                            $noteText = Normalize-Text $noteShape.TextFrame.TextRange.Text
                            if ($noteText.Length -gt 0 -and $noteText -notmatch "^Slide \d+$") {
                                $lines.Add(("S{0:00} Notes {1}: {2}" -f $slide.SlideIndex, $noteShape.Name, $noteText))
                            }
                        }
                    }
                }
            } catch {}
        }
        [System.IO.File]::WriteAllLines($OutPath, $lines, [System.Text.Encoding]::UTF8)
    } finally {
        if ($pres -ne $null) {
            try { $pres.Close() } catch {}
        }
        if ($pp -ne $null -and $pp.Presentations.Count -eq 0) {
            try { $pp.Quit() } catch {}
        }
    }
}

$workDir = "C:\Users\namma\.claude\aveva_bdm_resume_work"
$baseResume = "D:\이력서\AVEVA - Marine Principal Technical Support & Consultant – PLM SME, Busan\Joonha_Lee_AVEVA_Marine_PLM_SME_Resume.docx"
$inputResume = "D:\이력서\AVEVA - Business Development Manager\AVEVA_BDM_Resume_Codex_Input.docx"
$deck = "D:\이력서\AVEVA - Marine Principal Technical Support & Consultant – PLM SME, Busan\Proposal\Future_Industrial_PLM_Meeting_Deck_EN_V18.pptx"

Dump-WordDocument -Path $baseResume -OutPath (Join-Path $workDir "base_resume_dump.txt")
Dump-WordDocument -Path $inputResume -OutPath (Join-Path $workDir "bdm_input_dump.txt")
Dump-PowerPoint -Path $deck -OutPath (Join-Path $workDir "deck_v18_dump.txt")

"Dump complete: $workDir"
