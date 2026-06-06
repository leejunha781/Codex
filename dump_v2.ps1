$ErrorActionPreference = "Stop"
$path = "D:\이력서\ITT Cannon\Joonha_Lee_ITT_Cannon_FAE_Resume_Final_Integrated_v2.docx"
$out  = "C:\Users\namma\.claude\itt_work\dump_v2.txt"
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$lines = @()
try {
    $doc = $word.Documents.Open($path, $false, $true)  # ConfirmConversions=F, ReadOnly=T
    $lines += "===== PAGES: $($doc.ComputeStatistics(2))  WORDS: $($doc.ComputeStatistics(0)) ====="
    $lines += ""
    $lines += "##### PARAGRAPHS (in document order) #####"
    $i = 0
    foreach ($p in $doc.Paragraphs) {
        $i++
        $t = ($p.Range.Text -replace "[\r\n\a\x07]","").TrimEnd()
        $sz = $p.Range.Font.Size
        $bold = $p.Range.Font.Bold
        $styleName = ""
        try { $styleName = $p.Style.NameLocal } catch {}
        $inTable = $p.Range.Information(12)  # wdWithInTable
        if ($t.Trim().Length -eq 0 -and -not $inTable) {
            $lines += ("[{0}] (empty) sz={1} bold={2} style={3}" -f $i,$sz,$bold,$styleName)
        } elseif (-not $inTable) {
            $lines += ("[{0}] sz={1} bold={2} style=`"{3}`" :: {4}" -f $i,$sz,$bold,$styleName,$t)
        }
    }
    $lines += ""
    $lines += "##### TABLES #####"
    $tn = 0
    foreach ($tb in $doc.Tables) {
        $tn++
        $lines += "----- TABLE $tn : rows=$($tb.Rows.Count) cols=$($tb.Columns.Count) -----"
        $rn = 0
        foreach ($row in $tb.Rows) {
            $rn++
            $cellTexts = @()
            foreach ($cell in $row.Cells) {
                $ct = ($cell.Range.Text -replace "[\r\n\a\x07]"," ").Trim()
                $cellTexts += $ct
            }
            $lines += ("  R{0}: {1}" -f $rn, ($cellTexts -join "  |  "))
        }
    }
    $doc.Close([ref]$false)
} finally {
    if ($word.Documents.Count -eq 0) { $word.Quit() }
}
$lines | Out-File -FilePath $out -Encoding utf8
Write-Output "DONE -> $out"
