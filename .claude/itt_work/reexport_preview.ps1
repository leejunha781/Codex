$ErrorActionPreference = 'Stop'
$docx = "D:\이력서\ITT Cannon\Joonha_Lee_ITT_Cannon_FAE_Resume_Final_Integrated_v2.docx"
$pdf  = "C:\Users\namma\.claude\itt_work\v2_recolored.pdf"
if (Test-Path $pdf) { Remove-Item -LiteralPath $pdf -Force }
$word = New-Object -ComObject Word.Application
$word.Visible = $false
try { $word.DisplayAlerts = 0 } catch {}
$doc = $word.Documents.Open($docx, $false, $true)   # read-only open, just to export
$doc.ExportAsFixedFormat($pdf, 17)
$doc.Close([ref]$false)
if ($word.Documents.Count -eq 0) { $word.Quit() }
[Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
Write-Output ("PREVIEW_DONE size=" + (Get-Item -LiteralPath $pdf).Length)
