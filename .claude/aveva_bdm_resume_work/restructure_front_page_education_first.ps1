$ErrorActionPreference = "Stop"

function Normalize-Text {
    param([string]$Text)
    if ($null -eq $Text) { return "" }
    return (($Text -replace "[`r`n`a\x07]", " ") -replace "\s+", " ").Trim()
}

function Find-ParagraphExact {
    param(
        $Doc,
        [string]$Text
    )
    $needle = Normalize-Text $Text
    foreach ($p in $Doc.Paragraphs) {
        if ((Normalize-Text $p.Range.Text) -eq $needle) {
            return $p
        }
    }
    return $null
}

function Move-BlockBeforeHeading {
    param(
        $Doc,
        [string]$StartHeading,
        [string]$EndHeading,
        [string]$TargetHeading
    )
    $start = Find-ParagraphExact -Doc $Doc -Text $StartHeading
    $end = Find-ParagraphExact -Doc $Doc -Text $EndHeading
    $target = Find-ParagraphExact -Doc $Doc -Text $TargetHeading
    if ($null -eq $start -or $null -eq $end -or $null -eq $target) {
        throw "Could not find headings for moving block: $StartHeading / $EndHeading / $TargetHeading"
    }
    if ($start.Range.Start -ge $end.Range.Start) {
        throw "Invalid block order for $StartHeading to $EndHeading"
    }

    $block = $Doc.Range($start.Range.Start, $end.Range.Start)
    $block.Cut() | Out-Null

    $targetAfterCut = Find-ParagraphExact -Doc $Doc -Text $TargetHeading
    if ($null -eq $targetAfterCut) {
        throw "Target heading disappeared after cut: $TargetHeading"
    }
    $pasteRange = $Doc.Range($targetAfterCut.Range.Start, $targetAfterCut.Range.Start)
    $pasteRange.Collapse(1)
    $pasteRange.Paste()
}

function Tighten-TableByHeading {
    param(
        $Doc,
        [string]$Heading,
        [single]$FontSize
    )
    $headingPara = Find-ParagraphExact -Doc $Doc -Text $Heading
    if ($null -eq $headingPara) { return }

    $candidate = $null
    foreach ($t in $Doc.Tables) {
        if ($t.Range.Start -gt $headingPara.Range.Start) {
            if ($null -eq $candidate -or $t.Range.Start -lt $candidate.Range.Start) {
                $candidate = $t
            }
        }
    }
    if ($null -eq $candidate) { return }

    foreach ($cell in $candidate.Range.Cells) {
        $cell.Range.Font.Name = "Arial"
        $cell.Range.Font.Size = [single]$FontSize
        $cell.Range.ParagraphFormat.SpaceBefore = [single]0
        $cell.Range.ParagraphFormat.SpaceAfter = [single]0
        $cell.Range.ParagraphFormat.LineSpacingRule = 0
    }
    foreach ($row in $candidate.Rows) {
        try { $row.AllowBreakAcrossPages = 0 } catch {}
    }
}

function Tighten-EducationBlock {
    param($Doc)
    $education = Find-ParagraphExact -Doc $Doc -Text "Education"
    $career = Find-ParagraphExact -Doc $Doc -Text "Career Summary"
    if ($null -eq $education -or $null -eq $career) { return }

    foreach ($p in $Doc.Paragraphs) {
        if ($p.Range.Start -gt $education.Range.Start -and $p.Range.Start -lt $career.Range.Start) {
            $txt = Normalize-Text $p.Range.Text
            if ($txt.Length -gt 0) {
                $p.Range.Font.Name = "Arial"
                $p.Range.Font.Size = [single]10.5
                $p.Range.ParagraphFormat.SpaceBefore = [single]0
                $p.Range.ParagraphFormat.SpaceAfter = [single]0
                $p.Range.ParagraphFormat.LineSpacingRule = 0
            }
        }
    }
}

$docx = "D:\이력서\AVEVA - Business Development Manager\Joonha_Lee_AVEVA_BDM_Engineering_Marine_Korea_Japan_Seoul_Resume.docx"
$pdf = "D:\이력서\AVEVA - Business Development Manager\Joonha_Lee_AVEVA_BDM_Engineering_Marine_Korea_Japan_Seoul_Resume.pdf"

$word = New-Object -ComObject Word.Application
$word.Visible = $false
$doc = $null
try {
    $doc = $word.Documents.Open($docx, $false, $false)

    # Desired reviewer-first flow:
    # Basic Information -> Education -> Career Summary -> Job-Fit Map -> Core Competencies -> Detailed Experience.
    # First move Core Competencies behind Job-Fit/Education, then pull Education up to the first page.
    $core = Find-ParagraphExact -Doc $doc -Text "Core Competencies"
    $career = Find-ParagraphExact -Doc $doc -Text "Career Summary"
    $detailed = Find-ParagraphExact -Doc $doc -Text "Detailed Professional Experience"
    if ($core -ne $null -and $career -ne $null -and $detailed -ne $null -and $core.Range.Start -lt $career.Range.Start) {
        Move-BlockBeforeHeading -Doc $doc -StartHeading "Core Competencies" -EndHeading "Career Summary" -TargetHeading "Detailed Professional Experience"
    }

    $education = Find-ParagraphExact -Doc $doc -Text "Education"
    $career = Find-ParagraphExact -Doc $doc -Text "Career Summary"
    $core = Find-ParagraphExact -Doc $doc -Text "Core Competencies"
    if ($education -ne $null -and $career -ne $null -and $core -ne $null -and $education.Range.Start -gt $career.Range.Start) {
        Move-BlockBeforeHeading -Doc $doc -StartHeading "Education" -EndHeading "Core Competencies" -TargetHeading "Career Summary"
    }

    Tighten-EducationBlock -Doc $doc
    Tighten-TableByHeading -Doc $doc -Heading "Career Summary" -FontSize ([single]10.0)
    Tighten-TableByHeading -Doc $doc -Heading "Job-Fit Map: AVEVA Marine BDM Requirements vs. Career Evidence" -FontSize ([single]9.7)
    Tighten-TableByHeading -Doc $doc -Heading "Core Competencies" -FontSize ([single]10.0)

    $doc.Fields.Update() | Out-Null
    $doc.Save()
    $doc.ExportAsFixedFormat($pdf, 17)
    $pages = $doc.ComputeStatistics(2)
    $words = $doc.ComputeStatistics(0)
    $doc.Close([ref]$true)
    $doc = $null

    "UPDATED: $docx"
    "PDF: $pdf"
    "PAGES: $pages"
    "WORDS: $words"
} finally {
    if ($doc -ne $null) {
        try { $doc.Close([ref]$false) } catch {}
    }
    if ($word -ne $null -and $word.Documents.Count -eq 0) {
        try { $word.Quit() } catch {}
    }
}
