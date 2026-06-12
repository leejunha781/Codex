$ErrorActionPreference = "Stop"

$mainDocx = "D:\이력서\AVEVA - Business Development Manager\Joonha_Lee_AVEVA_BDM_Engineering_Marine_Korea_Japan_Seoul_Resume.docx"
$mainPdf = "D:\이력서\AVEVA - Business Development Manager\Joonha_Lee_AVEVA_BDM_Engineering_Marine_Korea_Japan_Seoul_Resume.pdf"
$docx = "D:\이력서\AVEVA - Business Development Manager\Joonha_Lee_AVEVA_BDM_Engineering_Marine_Korea_Japan_Seoul_Resume_Formal_Tone.docx"
$pdf = "D:\이력서\AVEVA - Business Development Manager\Joonha_Lee_AVEVA_BDM_Engineering_Marine_Korea_Japan_Seoul_Resume_Formal_Tone.pdf"
$tempPdf = "C:\Users\namma\.claude\aveva_bdm_resume_work\self_intro_no_blank_page.pdf"

if (Test-Path -LiteralPath $tempPdf) { Remove-Item -LiteralPath $tempPdf -Force }

$word = New-Object -ComObject Word.Application
$word.Visible = $false
$doc = $null
try {
    $doc = $word.Documents.Open($docx, $false, $false)

    $last = $doc.Paragraphs.Item($doc.Paragraphs.Count)
    $last.Range.Font.Size = [single]1
    $last.Range.Font.Hidden = [int]1
    $last.Range.ParagraphFormat.SpaceBefore = [single]0
    $last.Range.ParagraphFormat.SpaceAfter = [single]0
    $last.Range.ParagraphFormat.LineSpacingRule = 4
    $last.Range.ParagraphFormat.LineSpacing = [single]1

    $doc.Fields.Update() | Out-Null
    $doc.Save()
    $doc.ExportAsFixedFormat($tempPdf, 17)
    $pages = $doc.ComputeStatistics(2)
    $words = $doc.ComputeStatistics(0)
    $doc.Close([ref]$true)
    $doc = $null

    Copy-Item -LiteralPath $tempPdf -Destination $pdf -Force
    Copy-Item -LiteralPath $docx -Destination $mainDocx -Force
    Copy-Item -LiteralPath $pdf -Destination $mainPdf -Force

    "UPDATED_DOCX: $docx"
    "UPDATED_PDF: $pdf"
    "MAIN_DOCX: UPDATED"
    "MAIN_PDF: UPDATED"
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
