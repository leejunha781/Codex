$ErrorActionPreference = 'Continue'
$logPath = "C:\Users\namma\.claude\itt_work\recolor_log.txt"
$log = New-Object System.Text.StringBuilder
function L($m){ [void]$log.AppendLine($m); Set-Content -Path $logPath -Value $log.ToString() -Encoding UTF8 }

$path = "D:\이력서\ITT Cannon\Joonha_Lee_ITT_Cannon_FAE_Resume_Final_Integrated_v2.docx"
$pdf  = "D:\이력서\ITT Cannon\Joonha_Lee_ITT_Cannon_FAE_Resume_Final_Integrated_v2.pdf"

$charcoal  = 4210752    # RGB(64,64,64) - matches the heading underline
$lightgray = 15658734   # RGB(238,238,238)
$navy      = 7949855    # existing header navy
$lb1=16247773; $lb2=16247513; $lb3=16512494   # existing light-blue variants

$word = New-Object -ComObject Word.Application
$word.Visible = $false
try { $word.DisplayAlerts = 0 } catch {}
$doc = $word.Documents.Open($path, $false, $false)
L('opened')
try {
  $hd=0;$lg=0;$ft=0
  foreach($t in $doc.Tables){
    foreach($c in $t.Range.Cells){
      $bg=$null; try{$bg=$c.Shading.BackgroundPatternColor}catch{}
      if($bg -eq $navy){ try{$c.Shading.BackgroundPatternColor=[int]$charcoal; $hd++}catch{} }
      elseif($bg -eq $lb1 -or $bg -eq $lb2 -or $bg -eq $lb3){ try{$c.Shading.BackgroundPatternColor=[int]$lightgray; $lg++}catch{} }
      $fc=$null; try{$fc=$c.Range.Font.Color}catch{}
      if($fc -eq $navy){ try{$c.Range.Font.Color=[int]$charcoal; $ft++}catch{} }
    }
  }
  L("header-fills->charcoal=$hd  light-fills->gray=$lg  navy-text->charcoal=$ft")
  $doc.Save(); L('saved')
  try{ $doc.ExportAsFixedFormat($pdf,17); L('pdf') }catch{ L('pdf fail: '+$_.Exception.Message) }

  # verify remaining colored cells
  L('--- remaining fills ---')
  $ti=0
  foreach($t in $doc.Tables){ $ti++; foreach($c in $t.Range.Cells){ $bg=$null;try{$bg=$c.Shading.BackgroundPatternColor}catch{}; if($bg -ne $null -and $bg -ne -16777216 -and $bg -ne 16777215 -and $bg -ne $charcoal -and $bg -ne $lightgray){ L("  leftover T$ti bg=$bg") } } }
  L('done-scan')
}
catch { L('FATAL: '+$_.Exception.Message) }
finally {
  try{ $doc.Close([ref]$true) }catch{}
  try{ if($word.Documents.Count -eq 0){$word.Quit()} }catch{}
  try{ [Runtime.InteropServices.Marshal]::ReleaseComObject($word)|Out-Null }catch{}
}
Write-Output 'RECOLOR_DONE'
