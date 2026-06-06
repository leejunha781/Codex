$ErrorActionPreference = 'Continue'
$logPath = "C:\Users\namma\.claude\itt_work\headings_log.txt"
$log = New-Object System.Text.StringBuilder
function L($m){ [void]$log.AppendLine($m); Set-Content -Path $logPath -Value $log.ToString() -Encoding UTF8 }
function Norm($p){ return ($p.Range.Text -replace "[`r`n`a`f]","").Trim() }

$path = "D:\이력서\ITT Cannon\Joonha_Lee_ITT_Cannon_FAE_Resume_Final_Integrated_v2.docx"
$pdf  = "D:\이력서\ITT Cannon\Joonha_Lee_ITT_Cannon_FAE_Resume_Final_Integrated_v2.pdf"

$word = New-Object -ComObject Word.Application
$word.Visible = $false; try{$word.DisplayAlerts=0}catch{}
$doc = $word.Documents.Open($path,$false,$false)
L("opened pages=" + $doc.ComputeStatistics(2))

$fixes = @{
  'Professional Profile' = ' and Interest in ITT Cannon FAE Role'
  'Key Strengths'        = ' Matched to ITT Cannon FAE Role'
}

try {
  foreach($p in $doc.Paragraphs){
    $t = Norm $p
    $sz=$null;try{$sz=$p.Range.Font.Size}catch{}
    $bd=$null;try{$bd=$p.Range.Font.Bold}catch{}
    if($sz -eq 14 -and $bd -eq -1 -and $fixes.ContainsKey($t)){
      $append = $fixes[$t]
      $insertPos = $p.Range.End - 1
      $ir = $doc.Range($insertPos, $insertPos)
      $ir.InsertBefore($append)
      L("FIXED: '$t' + '$append'")
    }
  }

  # verify
  L("--- headings after fix ---")
  foreach($p in $doc.Paragraphs){
    $sz=$null;try{$sz=$p.Range.Font.Size}catch{}; $bd=$null;try{$bd=$p.Range.Font.Bold}catch{}
    if($sz -eq 14 -and $bd -eq -1 -and (Norm $p).Length -gt 0){ L("  H: " + (Norm $p)) }
  }

  $doc.Save(); L("saved")
  try{ $doc.ExportAsFixedFormat($pdf,17); L("pdf") }catch{ L("pdf fail: "+$_.Exception.Message) }
  L("FINAL pages=" + $doc.ComputeStatistics(2) + " words=" + $doc.ComputeStatistics(0))
}
catch { L("FATAL: "+$_.Exception.Message) }
finally {
  try{$doc.Close([ref]$true)}catch{}
  try{if($word.Documents.Count -eq 0){$word.Quit()}}catch{}
  try{[Runtime.InteropServices.Marshal]::ReleaseComObject($word)|Out-Null}catch{}
}
Write-Output "HEADINGS_DONE"
