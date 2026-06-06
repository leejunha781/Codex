$ErrorActionPreference = 'Continue'
$logPath = "C:\Users\namma\.claude\itt_work\fix_log.txt"
$log = New-Object System.Text.StringBuilder
function L($m){ [void]$log.AppendLine($m); Set-Content -Path $logPath -Value $log.ToString() -Encoding UTF8 }
function Norm($p){ return ($p.Range.Text -replace "[`r`n`a]","").Trim() }

$path = "D:\이력서\ITT Cannon\Joonha_Lee_ITT_Cannon_FAE_Resume_Final_Integrated_v2.docx"
$pdf  = "D:\이력서\ITT Cannon\Joonha_Lee_ITT_Cannon_FAE_Resume_Final_Integrated_v2.pdf"
$word = New-Object -ComObject Word.Application
$word.Visible = $false
try { $word.DisplayAlerts = 0 } catch {}
$doc = $word.Documents.Open($path, $false, $false)
L('opened')

try {
  # collect empty, size-14, bold paragraphs (stray heading-formatted blanks)
  $strays = New-Object System.Collections.ArrayList
  foreach($p in $doc.Paragraphs){
    $t = Norm $p
    if($t.Length -eq 0){
      $sz=$null;try{$sz=$p.Range.Font.Size}catch{}
      $bd=$null;try{$bd=$p.Range.Font.Bold}catch{}
      if($sz -eq 14 -and $bd -eq -1){ [void]$strays.Add($p) }
    }
  }
  L("strays found=$($strays.Count)")
  $del=0
  for($k=$strays.Count-1;$k -ge 0;$k--){
    $p=$strays[$k]
    try{ $p.Range.ParagraphFormat.Borders.Item([int]-3).LineStyle=[int]0 }catch{}
    try{ $rng=$p.Range.Duplicate; $rng.Delete() | Out-Null; $del++ }catch{ L('del fail: '+$_.Exception.Message) }
  }
  L("deleted=$del")

  $doc.Save(); L('saved')
  try{ $doc.ExportAsFixedFormat($pdf,17); L('pdf') }catch{ L('pdf fail: '+$_.Exception.Message) }

  L('PAGES='+$doc.ComputeStatistics(2))
  L('--- headings(14/bold) after cleanup ---')
  foreach($p in $doc.Paragraphs){ $sz=$null;try{$sz=$p.Range.Font.Size}catch{}; $bd=$null;try{$bd=$p.Range.Font.Bold}catch{}; if($sz -eq 14 -and $bd -eq -1){ L('  H: ['+(Norm $p)+']') } }
}
catch { L('FATAL: '+$_.Exception.Message) }
finally {
  try{ $doc.Close([ref]$true) }catch{}
  try{ if($word.Documents.Count -eq 0){$word.Quit()} }catch{}
  try{ [Runtime.InteropServices.Marshal]::ReleaseComObject($word)|Out-Null }catch{}
}
Write-Output 'FIX_DONE'
