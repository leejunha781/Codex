# Fix the doubled char in salary cell with exact trailing-marker counting
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
        $raw = $p.Range.Text
        if ((NormText $raw) -eq "회사 내규에 따라 협의의") {
            $tail = 0
            for ($i = $raw.Length - 1; $i -ge 0; $i--) {
                $c = [int]$raw[$i]
                if ($c -eq 13 -or $c -eq 7 -or $c -eq 10) { $tail++ } else { break }
            }
            $r = $p.Range.Duplicate
            $r.End = $r.End - $tail
            $r.Text = "회사 내규 협의"
            $fixed++
            Write-Output "FIXED salary cell (tail=$tail)"
        }
    }
    if ($fixed -ne 1) { Write-Output "WARN fixed=$fixed" }

    # verify final cell text + career summary line
    $tb = $doc.Tables.Item($doc.Tables.Count)
    foreach ($cell in $tb.Range.Cells) {
        if ($cell.RowIndex -eq 2 -and $cell.ColumnIndex -eq 1) {
            Write-Output ("r2c1 now: [{0}]" -f ($cell.Range.Text -replace "[\r\n\a\x07]","|"))
        }
    }
    foreach ($p in $doc.Paragraphs) {
        $n = NormText $p.Range.Text
        if ($n -like "*엑시큐어하이트론*" -and $n -like "*200*") { Write-Output ("career line: {0}" -f $n) }
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
