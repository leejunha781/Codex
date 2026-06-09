$ErrorActionPreference = "Stop"
$dst = "D:\이력서\모아컴코리아 - HW 개발\이준하_모아컴코리아_HW설계_최종_v2.docx"
$word = New-Object -ComObject Word.Application
$word.Visible = $false
try {
    $doc = $word.Documents.Open($dst, $false, $true)
    $tb = $doc.Tables.Item($doc.Tables.Count)   # salary table is the last one
    foreach ($cell in $tb.Range.Cells) {
        $ct = $cell.Range.Text
        $hex = ($ct.ToCharArray() | ForEach-Object { "{0:x2}" -f [int]$_ }) -join " "
        Write-Output ("r{0}c{1}: [{2}]" -f $cell.RowIndex, $cell.ColumnIndex, ($ct -replace "[\r\n\a\x07]","|"))
    }
    $doc.Close([ref]$false)
} finally {
    if ($word.Documents.Count -eq 0) { $word.Quit() }
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
}
