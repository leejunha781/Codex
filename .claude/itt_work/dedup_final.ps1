$ErrorActionPreference = 'Continue'
$logPath = "C:\Users\namma\.claude\itt_work\dedup_log.txt"
$log = New-Object System.Text.StringBuilder
function L($m){ [void]$log.AppendLine($m); Set-Content -Path $logPath -Value $log.ToString() -Encoding UTF8 }
function Norm($p){ return ($p.Range.Text -replace "[`r`n`a`f]","").Trim() }

$path = "D:\이력서\ITT Cannon\Joonha_Lee_ITT_Cannon_FAE_Resume_Final_Integrated_v2.docx"
$pdf  = "D:\이력서\ITT Cannon\Joonha_Lee_ITT_Cannon_FAE_Resume_Final_Integrated_v2.pdf"

$word = New-Object -ComObject Word.Application
$word.Visible = $false; try{$word.DisplayAlerts=0}catch{}
$doc = $word.Documents.Open($path,$false,$false)
L("opened pages=" + $doc.ComputeStatistics(2))

try {
  # ============================================================
  # STEP 1: List all connector-related headings and bullets
  # ============================================================
  L("--- connector/product section scan ---")
  foreach($p in $doc.Paragraphs){
    $t=Norm $p
    if($t -like '*Connector Product Knowledge*' -or $t -like '*Product Knowledge*Target*' -or
       $t -like '*Product/Application Knowledge*' -or $t -like 'Aerospace & Defense (preferred):*' -or
       $t -like '*connector families mapped*'){
      $pg=$null; try{$pg=$p.Range.Information(3)}catch{}
      $sz=$null;try{$sz=$p.Range.Font.Size}catch{}
      L("P$pg sz=$sz | " + $t.Substring(0,[Math]::Min(70,$t.Length)))
    }
  }

  # ============================================================
  # STEP 2: Remove OUR added connector section paragraphs
  # (detect by unique phrases from build_final.ps1 BC1-BC4 bullets)
  # The ORIGINAL section uses "Familiar with D38999-style"
  # OUR section uses "MIL-DTL-38999 / D38999 circular (Cannon KJL"
  # ============================================================
  $removeKeys = @(
    'ITT Cannon Connector Product Knowledge & Target-Market Fit',
    'ITT Cannon Product Knowledge & Target-Market Fit',
    'ITT Cannon connector families mapped to the Field Application Engineer'
  )
  # Bullets we added (start phrases unique to our version)
  $removeBulletPrefixes = @(
    'Aerospace & Defense (preferred): MIL-DTL-38999',
    'Transportation & Industrial (Rail, Heavy Vehicles, Energy): APD and APV',
    'Harsh-environment circular, electrical cabling, high-power / high-speed data',
    'Materials & manufacturing (high-performance thermoplastics'
  )

  $toRemove = New-Object System.Collections.ArrayList
  foreach($p in $doc.Paragraphs){
    $t=Norm $p
    $match=$false
    foreach($k in $removeKeys){ if($t -eq $k){ $match=$true; break } }
    if(-not $match){ foreach($k in $removeBulletPrefixes){ if($t.StartsWith($k)){ $match=$true; break } } }
    if($match){ [void]$toRemove.Add($p); L("MARK_REMOVE: "+$t.Substring(0,[Math]::Min(60,$t.Length))) }
  }
  $del=0
  for($k=$toRemove.Count-1;$k -ge 0;$k--){
    try{ $r=$toRemove[$k].Range.Duplicate; $r.Delete()|Out-Null; $del++ }catch{ L("del fail idx=$k") }
  }
  L("STEP2 removed=$del (expected 6 if full set present)")

  # ============================================================
  # STEP 3: Find Values paragraph (broader search)
  # ============================================================
  $vwFound=$false
  foreach($p in $doc.Paragraphs){
    $t=Norm $p
    if($t -like '*working habits*map*ITT*' -or $t -like '*Impeccable Character*Bold Thinking*' -or
       $t -like '*We Solve It*guiding principles*'){
      $pg=$null;try{$pg=$p.Range.Information(3)}catch{}
      L("VALUES_FOUND on page $pg | "+$t.Substring(0,[Math]::Min(70,$t.Length)))
      $vwFound=$true
    }
  }
  if(-not $vwFound){ L("VALUES_PARA: NOT found anywhere") }

  # ============================================================
  # STEP 4: Save + PDF + QA
  # ============================================================
  $doc.Save(); L("saved")
  try{ $doc.ExportAsFixedFormat($pdf,17); L("pdf") }catch{ L("pdf fail: "+$_.Exception.Message) }
  L("FINAL pages=" + $doc.ComputeStatistics(2) + " words=" + $doc.ComputeStatistics(0) + " tables=" + $doc.Tables.Count)

  # Verify connector sections that remain
  L("--- remaining connector headings ---")
  foreach($p in $doc.Paragraphs){
    $t=Norm $p
    $sz=$null;try{$sz=$p.Range.Font.Size}catch{}
    $bd=$null;try{$bd=$p.Range.Font.Bold}catch{}
    if($sz -eq 14 -and $bd -eq -1 -and $t.Length -gt 0){ L("  H14: "+$t) }
  }
}
catch { L("FATAL: "+$_.Exception.Message) }
finally {
  try{$doc.Close([ref]$true)}catch{}
  try{if($word.Documents.Count -eq 0){$word.Quit()}}catch{}
  try{[Runtime.InteropServices.Marshal]::ReleaseComObject($word)|Out-Null}catch{}
}
Write-Output "DEDUP_DONE"
