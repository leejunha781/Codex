# Extract interview-strategy doc (read-only)
$ErrorActionPreference = "Stop"
$src = "D:\이력서\모아컴코리아 - HW 개발\이준하_모아컴코리아_HW설계_전문면접_기술이론_전략_수식포함.docx"
$out = "C:\Users\namma\.claude\moacom_work\dump_interview.txt"

$word = New-Object -ComObject Word.Application
$word.Visible = $false
try {
    $doc = $word.Documents.Open($src, $false, $true)
    try {
        $sb = New-Object System.Text.StringBuilder
        [void]$sb.AppendLine("PAGES: $($doc.ComputeStatistics(2))  WORDS: $($doc.ComputeStatistics(0))  TABLES: $($doc.Tables.Count)  OMATHS: $($doc.OMaths.Count)")
        $i = 0
        foreach ($p in $doc.Paragraphs) {
            $i++
            $t = ($p.Range.Text -replace "[\r\n\a\x07]","" -replace "\s+"," ").Trim()
            $inTable = $p.Range.Information(12)
            $style = ""; try { $style = $p.Range.Style.NameLocal } catch {}
            $tag = if ($inTable) { "T" } else { "P" }
            [void]$sb.AppendLine(("[{0:d3}{1}] s={2} f={3}/{4} b={5} | {6}" -f $i, $tag, $style, $p.Range.Font.Name, $p.Range.Font.Size, $p.Range.Font.Bold, $t))
        }
        $ti = 0
        foreach ($tb in $doc.Tables) {
            $ti++
            [void]$sb.AppendLine("--- Table $ti : Rows=$($tb.Rows.Count) Cols=$($tb.Columns.Count)")
        }
        [System.IO.File]::WriteAllText($out, $sb.ToString(), [System.Text.Encoding]::UTF8)
        Write-Output "OK: $out"
    } finally { $doc.Close([ref]$false) }
} finally {
    if ($word.Documents.Count -eq 0) { $word.Quit() }
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
}
