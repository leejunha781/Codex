$ErrorActionPreference = "Stop"
$path = "D:\이력서\ITT Cannon\Joonha_Lee_ITT_Cannon_FAE_Resume_Final_Integrated_v2.docx"
$out  = "C:\Users\namma\.claude\itt_work\dump_full.txt"
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$lines = @()
try {
    $doc = $word.Documents.Open($path, $false, $true)
    # Header/footer tables and any merged tables: dump via merge-safe Range.Cells
    $tn = 0
    foreach ($tb in $doc.Tables) {
        $tn++
        $lines += "===== TABLE $tn : rows=$($tb.Rows.Count) cols=$($tb.Columns.Count) ====="
        $ci = 0
        foreach ($cell in $tb.Range.Cells) {
            $ci++
            $ct = ($cell.Range.Text -replace "[\r\n\a\x07]"," ").Trim()
            $lines += ("   cell[r$($cell.RowIndex),c$($cell.ColumnIndex)]: {0}" -f $ct)
        }
        $lines += ""
    }
    $doc.Close([ref]$false)
} finally {
    if ($word.Documents.Count -eq 0) { $word.Quit() }
}
$lines | Out-File -FilePath $out -Encoding utf8
Write-Output "DONE -> $out"
