$ErrorActionPreference = "Stop"

function Normalize-Text {
    param([string]$Text)
    if ($null -eq $Text) { return "" }
    return (($Text -replace "[`r`n`a\x07]", " ") -replace "\s+", " ").Trim()
}

$path = "D:\이력서\AVEVA - Business Development Manager\Joonha_Lee_AVEVA_BDM_Engineering_Marine_Korea_Japan_Seoul_Resume.docx"
$pdf = "D:\이력서\AVEVA - Business Development Manager\Joonha_Lee_AVEVA_BDM_Engineering_Marine_Korea_Japan_Seoul_Resume.pdf"

$word = New-Object -ComObject Word.Application
$word.Visible = $false
$doc = $null
try {
    $doc = $word.Documents.Open($path, $false, $false)

    # Remove the manual page break that originally forced Detailed Professional Experience to start on a fresh page.
    # After the BDM job-fit map became longer, that break created a mostly blank page.
    $breakChar = [string][char]12
    $toDelete = New-Object System.Collections.Generic.List[object]
    foreach ($p in $doc.Paragraphs) {
        $raw = $p.Range.Text
        if ($raw -like ("*" + $breakChar + "*")) {
            $toDelete.Add($p) | Out-Null
        }
    }
    for ($i = $toDelete.Count - 1; $i -ge 0; $i--) {
        try { $toDelete[$i].Range.Delete() | Out-Null } catch {}
    }

    # Compact the front alignment tables without changing the overall visual style.
    foreach ($idx in @(3,5)) {
        $tb = $doc.Tables.Item($idx)
        foreach ($cell in $tb.Range.Cells) {
            $cell.Range.Font.Size = [single]10
            $cell.Range.ParagraphFormat.SpaceBefore = [single]0
            $cell.Range.ParagraphFormat.SpaceAfter = [single]0
            $cell.Range.ParagraphFormat.LineSpacingRule = 0
        }
    }
    foreach ($cell in $doc.Tables.Item(4).Range.Cells) {
        $cell.Range.Font.Size = [single]10.5
        $cell.Range.ParagraphFormat.SpaceBefore = [single]0
        $cell.Range.ParagraphFormat.SpaceAfter = [single]0
        $cell.Range.ParagraphFormat.LineSpacingRule = 0
    }

    foreach ($p in $doc.Paragraphs) {
        $txt = Normalize-Text $p.Range.Text
        if ($txt -eq "Career Summary") {
            $p.Range.ParagraphFormat.KeepWithNext = [int]-1
        }
        if ($txt -eq "Detailed Professional Experience") {
            $p.Range.ParagraphFormat.KeepWithNext = [int]-1
        }
    }

    $doc.Fields.Update() | Out-Null
    $doc.Save()
    $doc.ExportAsFixedFormat($pdf, 17)
    $pages = $doc.ComputeStatistics(2)
    $words = $doc.ComputeStatistics(0)
    $doc.Close([ref]$true)
    $doc = $null
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
