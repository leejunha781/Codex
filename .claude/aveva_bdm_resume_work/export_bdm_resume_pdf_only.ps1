$ErrorActionPreference = "Stop"

$docx = "D:\이력서\AVEVA - Business Development Manager\Joonha_Lee_AVEVA_BDM_Engineering_Marine_Korea_Japan_Seoul_Resume.docx"
$pdf = "D:\이력서\AVEVA - Business Development Manager\Joonha_Lee_AVEVA_BDM_Engineering_Marine_Korea_Japan_Seoul_Resume.pdf"
$tempDocx = "C:\Users\namma\.claude\aveva_bdm_resume_work\pdf_export_source.docx"
$tempPdf = "C:\Users\namma\.claude\aveva_bdm_resume_work\pdf_export_source.pdf"

Copy-Item -LiteralPath $docx -Destination $tempDocx -Force
if (Test-Path -LiteralPath $tempPdf) { Remove-Item -LiteralPath $tempPdf -Force }

$word = New-Object -ComObject Word.Application
$word.Visible = $false
$doc = $null
try {
    # Export from a local temporary copy, then replace the final PDF.
    $doc = $word.Documents.Open($tempDocx, $false, $true)
    $doc.ExportAsFixedFormat($tempPdf, 17)
    $pages = $doc.ComputeStatistics(2)
    $words = $doc.ComputeStatistics(0)
    $doc.Close([ref]$false)
    $doc = $null
    Copy-Item -LiteralPath $tempPdf -Destination $pdf -Force
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
