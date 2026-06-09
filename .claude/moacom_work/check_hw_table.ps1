$ErrorActionPreference = "Stop"
$p = "D:\이력서\모아컴코리아 - HW 개발\이준하_모아컴코리아_HW 설계.docx"
$word = New-Object -ComObject Word.Application
$word.Visible = $false
try {
    $doc = $word.Documents.Open($p, $false, $true)
    Write-Output ("Tables: {0}" -f $doc.Tables.Count)
    $ti = 0
    foreach ($tb in $doc.Tables) {
        $ti++
        if ($ti -eq 1) { continue }   # skip the big header table
        Write-Output ("--- Table {0}: Rows={1} Cols={2}" -f $ti, $tb.Rows.Count, $tb.Columns.Count)
        foreach ($cell in $tb.Range.Cells) {
            $ct = ($cell.Range.Text -replace "[\r\n\a\x07]"," " -replace "\s+"," ").Trim()
            Write-Output ("  r{0}c{1}: {2}" -f $cell.RowIndex, $cell.ColumnIndex, $ct)
        }
    }
    # also show paragraphs around 연봉 정보 heading
    $show = $false
    foreach ($par in $doc.Paragraphs) {
        $t = ($par.Range.Text -replace "[\r\n\a\x07]","" -replace "\s+"," ").Trim()
        if ($t -like "*연봉 정보*") { $show = $true }
        if ($show -and $t.Length -gt 0) { Write-Output ("P: {0}" -f $t) }
        if ($show -and $t -like "*자기소개서*") { break }
    }
    $doc.Close([ref]$false)
} finally {
    if ($word.Documents.Count -eq 0) { $word.Quit() }
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
}
