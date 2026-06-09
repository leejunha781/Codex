# Two micro-fixes on the v2 file, then re-export PDF
$ErrorActionPreference = "Stop"
$dst = "D:\이력서\모아컴코리아 - HW 개발\이준하_모아컴코리아_HW설계_최종_v2.docx"
$pdf = "D:\이력서\모아컴코리아 - HW 개발\이준하_모아컴코리아_HW설계_최종_v2.pdf"
function NormText([string]$t) { ($t -replace "[\r\n\a\x07]","" -replace "\s+"," ").Trim() }

$word = New-Object -ComObject Word.Application
$word.Visible = $false
try {
    $doc = $word.Documents.Open($dst, $false, $false)
    $n1 = 0; $n2 = 0
    foreach ($p in $doc.Paragraphs) {
        $raw = $p.Range.Text
        $norm = NormText $raw
        if (($norm -like "*엑시큐어하이트론*") -and ($norm -like "*2007년 04월*")) {
            $r = $p.Range.Duplicate
            if ($raw -match "\x07$") { $r.End = $r.End - 2 } else { $r.End = $r.End - 1 }
            $r.Text = $r.Text.Replace("2007년 04월", "2007년 05월")
            $n1++; Write-Output "FIXED career-summary end month"
        }
        elseif ($norm -eq "회사 내규에 따라 협의") {
            $r = $p.Range.Duplicate
            if ($raw -match "\x07$") { $r.End = $r.End - 2 } else { $r.End = $r.End - 1 }
            $r.Text = "회사 내규 협의"
            $n2++; Write-Output "FIXED salary wording"
        }
    }
    if ($n1 -ne 1 -or $n2 -ne 1) { Write-Output "WARN counts n1=$n1 n2=$n2" }
    $doc.Save()
    Write-Output ("PAGES: {0}" -f $doc.ComputeStatistics(2))
    $doc.ExportAsFixedFormat($pdf, 17)
    Write-Output "PDF re-exported"
    $doc.Close([ref]$false)
} finally {
    if ($word.Documents.Count -eq 0) { $word.Quit() }
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
}
