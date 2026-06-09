# Compare text content of the three candidate files
$ErrorActionPreference = "Stop"
$files = @(
    @{ Tag="v2";      Path="D:\이력서\모아컴코리아 - HW 개발\이준하_모아컴코리아_HW설계_최종_v2.docx" },
    @{ Tag="hwsulgye";Path="D:\이력서\모아컴코리아 - HW 개발\이준하_모아컴코리아_HW 설계.docx" },
    @{ Tag="choejong";Path="D:\이력서\모아컴코리아 - HW 개발\이준하_모아컴코리아_HW설계_최종.docx" }
)
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$texts = @{}
try {
    foreach ($f in $files) {
        $doc = $word.Documents.Open($f.Path, $false, $true)
        try {
            $lines = @()
            foreach ($p in $doc.Paragraphs) {
                $t = ($p.Range.Text -replace "[\r\n\a\x07]","" -replace "\s+"," ").Trim()
                if ($t.Length -gt 0) { $lines += $t }
            }
            $texts[$f.Tag] = $lines
            $out = "C:\Users\namma\.claude\moacom_work\cmp_" + $f.Tag + ".txt"
            [System.IO.File]::WriteAllLines($out, $lines, [System.Text.Encoding]::UTF8)
            Write-Output ("{0}: paras={1} pages={2}" -f $f.Tag, $lines.Count, $doc.ComputeStatistics(2))
        } finally { $doc.Close([ref]$false) }
    }
} finally {
    if ($word.Documents.Count -eq 0) { $word.Quit() }
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
}
Write-Output "--- v2 vs HW설계 ---"
$d1 = Compare-Object $texts["v2"] $texts["hwsulgye"]
if ($d1) { $d1 | ForEach-Object { Write-Output ("{0} {1}" -f $_.SideIndicator, $_.InputObject.Substring(0,[Math]::Min(80,$_.InputObject.Length))) } } else { Write-Output "IDENTICAL" }
Write-Output "--- v2 vs 최종(8:26 saved) ---"
$d2 = Compare-Object $texts["v2"] $texts["choejong"]
if ($d2) { $d2 | ForEach-Object { Write-Output ("{0} {1}" -f $_.SideIndicator, $_.InputObject.Substring(0,[Math]::Min(80,$_.InputObject.Length))) } } else { Write-Output "IDENTICAL" }
