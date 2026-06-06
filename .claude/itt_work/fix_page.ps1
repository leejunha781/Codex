$ErrorActionPreference = 'Continue'
$logPath = "C:\Users\namma\.claude\itt_work\fix_page_log.txt"
$log = New-Object System.Text.StringBuilder
function L($m){ [void]$log.AppendLine($m); Set-Content -Path $logPath -Value $log.ToString() -Encoding UTF8 }
function Norm($p){ return ($p.Range.Text -replace "[`r`n`a`f]","").Trim() }

$path = "D:\이력서\ITT Cannon\Joonha_Lee_ITT_Cannon_FAE_Resume_Final_Integrated_v2.docx"
$pdf  = "D:\이력서\ITT Cannon\Joonha_Lee_ITT_Cannon_FAE_Resume_Final_Integrated_v2.pdf"

$word = New-Object -ComObject Word.Application
$word.Visible = $false; try{$word.DisplayAlerts=0}catch{}
$doc = $word.Documents.Open($path,$false,$false)
L("opened pages=" + $doc.ComputeStatistics(2) + " sections=" + $doc.Sections.Count)

try {
  # ============================================================
  # STEP 1: Diagnose blank-page cause
  #  - Report PageBreakBefore on all non-table paragraphs
  #  - Report section break markers
  # ============================================================
  $pbBef=0
  foreach($p in $doc.Paragraphs){
    $inT=$false; try{$inT=$p.Range.Information(12)}catch{}
    if($inT){continue}
    $pb=$false; try{$pb=$p.PageBreakBefore}catch{}
    if($pb){
      $pbBef++
      L("PageBreakBefore=T: [" + (Norm $p).Substring(0,[Math]::Min(40,(Norm $p).Length)) + "]")
    }
  }
  L("PageBreakBefore paragraphs=$pbBef")

  # ============================================================
  # STEP 2: Section breaks — remove extra sections by replacing
  # section break chars with nothing in each section's last para
  # ============================================================
  $sectCount=$doc.Sections.Count
  L("sections=$sectCount")
  if($sectCount -gt 1){
    for($si=1;$si -lt $sectCount;$si++){  # iterate all but last
      $sect=$doc.Sections.Item($si)
      $lastParaRange=$sect.Range.Duplicate
      # Section end character is at $sect.Range.End - 1
      # Replace the section break by modifying the range
      try{
        $sr=$doc.Range($sect.Range.End-1,$sect.Range.End-1)
        $sr.Delete()|Out-Null
        L("removed section break $si")
      }catch{ L("sect break $si remove fail: "+$_.Exception.Message) }
    }
  }

  # ============================================================
  # STEP 3: Reset PageBreakBefore=False on all non-table paras
  # ============================================================
  $pbReset=0
  foreach($p in $doc.Paragraphs){
    $inT=$false; try{$inT=$p.Range.Information(12)}catch{}
    if($inT){continue}
    $pb=$false; try{$pb=$p.PageBreakBefore}catch{}
    if($pb){ try{$p.PageBreakBefore=$false; $pbReset++}catch{} }
  }
  L("STEP3 PageBreakBefore reset=$pbReset")

  # ============================================================
  # STEP 4: Show current Key Strength bullets to find B5
  # ============================================================
  $inKS=$false; $bi=0
  foreach($p in $doc.Paragraphs){
    $t=Norm $p
    if($t -eq 'Key Strengths Matched to ITT Cannon FAE Role'){$inKS=$true;continue}
    if($inKS){
      $st='';try{$st=$p.Style.NameLocal}catch{}
      if($st -eq '글머리 기호' -or $st -eq 'List Bullet'){
        $bi++; $bLabel="KS_bullet"+$bi; L("${bLabel}: [" + $t.Substring(0,[Math]::Min(60,$t.Length)) + "]")
      } elseif($t.Length -gt 0 -and $bi -gt 0){ break }  # stop at next non-bullet
    }
  }

  # ============================================================
  # STEP 5: Extend B5 by matching via partial text
  # ============================================================
  $b5done=$false
  foreach($p in $doc.Paragraphs){
    $t=Norm $p
    if($t -like 'New Design*'){
      L("B5 candidate: [" + $t.Substring(0,[Math]::Min(60,$t.Length)) + "]")
      try{
        $insertPos=$p.Range.End-1
        $ir=$doc.Range($insertPos,$insertPos)
        $ir.InsertBefore(' Hands-on hardware design experience - ARM Cortex-M3 CPU/I/O boards, sensor and communication interfaces, embedded control circuits - provides practical insight for customer discussions on new component and assembly design-in.')
        $b5done=$true; L("B5 extended OK")
      }catch{ L("B5 extend fail: "+$_.Exception.Message) }
      break
    }
  }
  if(-not $b5done){ L("B5 NOT FOUND via 'New Design*'") }

  # ============================================================
  # STEP 6: Verify Values paragraph
  # ============================================================
  $vw=$false
  foreach($p in $doc.Paragraphs){
    $t=Norm $p
    if($t -like 'These working habits*'){ $vw=$true; L("Values para FOUND: ["+$t.Substring(0,[Math]::Min(50,$t.Length))+"]"); break }
    if($t -like '*also map to ITT*'){ $vw=$true; L("Values para FOUND (mid-match): ["+$t.Substring(0,[Math]::Min(50,$t.Length))+"]"); break }
  }
  if(-not $vw){ L("Values para NOT FOUND") }

  # ============================================================
  # STEP 7: Save + PDF + final page count
  # ============================================================
  $doc.Save(); L("saved")
  try{ $doc.ExportAsFixedFormat($pdf,17); L("pdf") }catch{ L("pdf fail: "+$_.Exception.Message) }
  L("FINAL pages=" + $doc.ComputeStatistics(2) + " words=" + $doc.ComputeStatistics(0))
}
catch{ L("FATAL: "+$_.Exception.Message) }
finally{
  try{$doc.Close([ref]$true)}catch{}
  try{if($word.Documents.Count -eq 0){$word.Quit()}}catch{}
  try{[Runtime.InteropServices.Marshal]::ReleaseComObject($word)|Out-Null}catch{}
}
Write-Output "FIX_PAGE_DONE"
