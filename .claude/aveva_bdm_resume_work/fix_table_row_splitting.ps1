$ErrorActionPreference = "Stop"

$docx = "D:\이력서\AVEVA - Business Development Manager\Joonha_Lee_AVEVA_BDM_Engineering_Marine_Korea_Japan_Seoul_Resume.docx"
$pdf = "D:\이력서\AVEVA - Business Development Manager\Joonha_Lee_AVEVA_BDM_Engineering_Marine_Korea_Japan_Seoul_Resume.pdf"

$word = New-Object -ComObject Word.Application
$word.Visible = $false
$doc = $null
try {
    $doc = $word.Documents.Open($docx, $false, $false)

    # Keep summary and JD-map rows intact when they meet a page boundary.
    foreach ($idx in @(4,5)) {
        $tb = $doc.Tables.Item($idx)
        foreach ($row in $tb.Rows) {
            try { $row.AllowBreakAcrossPages = [int]0 } catch {}
        }
    }

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
