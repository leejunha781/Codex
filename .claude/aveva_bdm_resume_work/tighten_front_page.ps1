$ErrorActionPreference = "Stop"

function Set-RangeTextKeepMark {
    param($Range, [string]$Text)
    $r = $Range.Duplicate
    if ($r.End -gt $r.Start) { $r.End = $r.End - 1 }
    $r.Text = $Text
}

$docx = "D:\이력서\AVEVA - Business Development Manager\Joonha_Lee_AVEVA_BDM_Engineering_Marine_Korea_Japan_Seoul_Resume.docx"
$pdf = "D:\이력서\AVEVA - Business Development Manager\Joonha_Lee_AVEVA_BDM_Engineering_Marine_Korea_Japan_Seoul_Resume.pdf"

$word = New-Object -ComObject Word.Application
$word.Visible = $false
$doc = $null
try {
    $doc = $word.Documents.Open($docx, $false, $false)
    $cell = $doc.Tables.Item(4).Range.Cells.Item(10)
    Set-RangeTextKeepMark -Range $cell.Range -Text "Global customer engagement; requirements/evidence control, 400+ verification items and release readiness."
    $cell.Range.Font.Name = "Arial"
    $cell.Range.Font.Size = [single]10.5
    $cell.Range.Font.Bold = [int]0
    $cell.Range.ParagraphFormat.SpaceBefore = [single]0
    $cell.Range.ParagraphFormat.SpaceAfter = [single]0
    $cell.Range.ParagraphFormat.LineSpacingRule = 0
    $doc.Save()
    $doc.ExportAsFixedFormat($pdf, 17)
    $pages = $doc.ComputeStatistics(2)
    $doc.Close([ref]$true)
    $doc = $null
    "PAGES: $pages"
} finally {
    if ($doc -ne $null) {
        try { $doc.Close([ref]$false) } catch {}
    }
    if ($word -ne $null -and $word.Documents.Count -eq 0) {
        try { $word.Quit() } catch {}
    }
}
