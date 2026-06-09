# Cell marker = ONE story position (though Text shows \r\a): use End-1
$ErrorActionPreference = "Stop"
$dst = "D:\이력서\모아컴코리아 - HW 개발\이준하_모아컴코리아_HW설계_최종_v2.docx"
$pdf = "D:\이력서\모아컴코리아 - HW 개발\이준하_모아컴코리아_HW설계_최종_v2.pdf"
function NormText([string]$t) { ($t -replace "[\r\n\a\x07]","" -replace "\s+"," ").Trim() }

$word = New-Object -ComObject Word.Application
$word.Visible = $false
try {
    $doc = $word.Documents.Open($dst, $false, $false)
    $fixed = 0
    foreach ($p in $doc.Paragraphs) {
        if ((NormText $p.Range.Text) -eq "회사 내규 협의의") {
            $r = $p.Range.Duplicate
            $r.End = $r.End - 1
            $r.Text = "회사 내규 협의"
            $fixed++
        }
    }
    Write-Output "fixed=$fixed"
    $tb = $doc.Tables.Item($doc.Tables.Count)
    foreach ($cell in $tb.Range.Cells) {
        if ($cell.RowIndex -eq 2 -and $cell.ColumnIndex -eq 1) {
            Write-Output ("r2c1 now: [{0}]" -f ($cell.Range.Text -replace "[\r\n\a\x07]","|"))
        }
    }
    $doc.Save()
    Write-Output ("PAGES: {0}" -f $doc.ComputeStatistics(2))
    $doc.ExportAsFixedFormat($pdf, 17)
    Write-Output "PDF re-exported"
    $doc.Close([ref]$false)
} finally {
    if ($word.Documents.Count -eq 0) { $word.Quit() }
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
}
