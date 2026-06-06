$ErrorActionPreference = "Stop"
$path = "D:\이력서\ITT Cannon\Joonha_Lee_ITT_Cannon_FAE_Resume_Final_Integrated_v2.docx"
$word = New-Object -ComObject Word.Application
$word.Visible = $false
try {
    $doc = $word.Documents.Open($path, $false, $true)
    $i = 0
    foreach ($p in $doc.Paragraphs) {
        $i++
        $raw = ($p.Range.Text -replace "[\r\n\a\x07]","")
        if ($raw -match "nowledge" -or $raw -match "pplication" -or $raw -match "ntegrated" -or $raw -match "lignment" -or $raw -match "Aerospace & Defense:" -or $raw -match "Education, Cert") {
            $codes = ($raw.ToCharArray() | Select-Object -First 40 | ForEach-Object { [int][char]$_ }) -join ","
            Write-Output ("[{0}] LEN={1} TEXT=<{2}>" -f $i, $raw.Length, $raw)
            Write-Output ("       firstcodes: {0}" -f $codes)
        }
    }
    $doc.Close([ref]$false)
} finally {
    if ($word.Documents.Count -eq 0) { $word.Quit() }
}
Write-Output "DIAG DONE"
