$ErrorActionPreference = 'Stop'
$word = New-Object -ComObject Word.Application
$word.Visible = $false
try { $word.DisplayAlerts = 0 } catch {}

function Extract-Doc($path, $outText, $outHi) {
    $doc = $word.Documents.Open($path, $false, $true)  # ConfirmConversions=$false, ReadOnly=$true
    $txt = $doc.Content.Text
    Set-Content -Path $outText -Value $txt -Encoding UTF8

    # Detect highlighted runs (any non-zero HighlightColorIndex)
    $sb = New-Object System.Text.StringBuilder
    $cur = ""
    $curColor = -1
    foreach ($w in $doc.Words) {
        $hc = 0
        try { $hc = $w.HighlightColorIndex } catch { $hc = 0 }
        if ($hc -ne 0) {
            $cur += $w.Text
            $curColor = $hc
        } else {
            if ($cur.Trim().Length -gt 0) {
                [void]$sb.AppendLine("[color=$curColor] " + ($cur -replace '\s+',' ').Trim())
            }
            $cur = ""
            $curColor = -1
        }
    }
    if ($cur.Trim().Length -gt 0) {
        [void]$sb.AppendLine("[color=$curColor] " + ($cur -replace '\s+',' ').Trim())
    }
    Set-Content -Path $outHi -Value $sb.ToString() -Encoding UTF8

    $doc.Close([ref]$false)
}

Extract-Doc "D:\이력서\ITT Cannon\Adecco Korea_Field Application Engineer (FAE).pdf" "C:\Users\namma\.claude\itt_work\jd_text.txt" "C:\Users\namma\.claude\itt_work\jd_highlights.txt"
Extract-Doc "D:\이력서\ITT Cannon\Joonha_Lee_ITT_Cannon_FAE_Resume_Standard_Format.docx" "C:\Users\namma\.claude\itt_work\template_text.txt" "C:\Users\namma\.claude\itt_work\template_highlights.txt"

if ($word.Documents.Count -eq 0) { $word.Quit() }
[Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
Write-Output "EXTRACTION_DONE"
