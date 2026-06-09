# Extract text + structure from both MoaComm resume docs (read-only)
$ErrorActionPreference = "Stop"
$files = @( @{ Path = "D:\이력서\모아컴코리아 - HW 개발\이준하_모아컴코리아_HW설계_최종_v2.docx"; Out = "C:\Users\namma\.claude\moacom_work\dump_v2.txt" } )

$word = New-Object -ComObject Word.Application
$word.Visible = $false
try {
    foreach ($f in $files) {
        $doc = $word.Documents.Open($f.Path, $false, $true)  # ConfirmConversions=F, ReadOnly=T
        try {
            $sb = New-Object System.Text.StringBuilder
            [void]$sb.AppendLine("=== FILE: $($f.Path)")
            [void]$sb.AppendLine("PAGES: $($doc.ComputeStatistics(2))  WORDS: $($doc.ComputeStatistics(0))  TABLES: $($doc.Tables.Count)")
            [void]$sb.AppendLine("")
            $i = 0
            foreach ($p in $doc.Paragraphs) {
                $i++
                $t = ($p.Range.Text -replace "[\r\n\a\x07]","").Trim()
                $inTable = $p.Range.Information(12)  # wdWithInTable
                $style = ""
                try { $style = $p.Range.Style.NameLocal } catch {}
                $fn = $p.Range.Font.Name
                $fs = $p.Range.Font.Size
                $fb = $p.Range.Font.Bold
                $ls = $p.Format.LineSpacingRule
                $sa = $p.Format.SpaceAfter
                $tag = if ($inTable) { "T" } else { "P" }
                [void]$sb.AppendLine(("[{0:d3}{1}] style={2} font={3}/{4} bold={5} lsr={6} sa={7} | {8}" -f $i, $tag, $style, $fn, $fs, $fb, $ls, $sa, $t))
            }
            [void]$sb.AppendLine("")
            [void]$sb.AppendLine("=== TABLES ===")
            $ti = 0
            foreach ($tb in $doc.Tables) {
                $ti++
                [void]$sb.AppendLine("--- Table $ti : Rows=$($tb.Rows.Count) Cols=$($tb.Columns.Count)")
                foreach ($cell in $tb.Range.Cells) {
                    $ct = ($cell.Range.Text -replace "[\r\n\a\x07]"," " -replace "\s+"," ").Trim()
                    [void]$sb.AppendLine(("  r{0}c{1}: {2}" -f $cell.RowIndex, $cell.ColumnIndex, $ct))
                }
            }
            [System.IO.File]::WriteAllText($f.Out, $sb.ToString(), [System.Text.Encoding]::UTF8)
            Write-Output "OK: $($f.Out)"
        } finally {
            $doc.Close([ref]$false)
        }
    }
} finally {
    if ($word.Documents.Count -eq 0) { $word.Quit() }
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
}
