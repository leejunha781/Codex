$ErrorActionPreference = "Stop"
Get-Process WINWORD -ErrorAction SilentlyContinue | Stop-Process -Force
Start-Sleep -Milliseconds 600
foreach($n in @("StartupItems","DocumentRecovery")){
  $k = "HKCU:\Software\Microsoft\Office\16.0\Word\Resiliency\$n"
  if (Test-Path $k) { Remove-Item $k -Recurse -Force }
}
$null = Start-Process "WINWORD.EXE" -PassThru
Start-Sleep -Seconds 5
$warm = [bool](Get-Process WINWORD -ErrorAction SilentlyContinue)
Write-Output ("warm=" + $warm)

$test = "C:\Users\namma\.claude\plm_slide_work\word_com_test.docx"
$pdf  = [System.IO.Path]::ChangeExtension($test, ".pdf")
$w = $null
for ($i=0; $i -lt 6 -and $w -eq $null; $i++){ try { $w = New-Object -ComObject Word.Application } catch { Start-Sleep -Milliseconds 1200 } }
if ($w -eq $null) { throw "Word COM did not start" }
$w.Visible = $false
$w.DisplayAlerts = 0
$d = $w.Documents.Add()
$d.Content.Text = "Word COM test OK"
$d.SaveAs([ref]$test, [ref]16)
$d.ExportAsFixedFormat($pdf, 17)
$d.Close([ref]$false)
$w.Quit()
Write-Output ("docx=" + (Test-Path $test) + " pdf=" + (Test-Path $pdf))
