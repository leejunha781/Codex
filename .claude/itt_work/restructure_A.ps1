$ErrorActionPreference = 'Continue'
$logPath = "C:\Users\namma\.claude\itt_work\A_log.txt"
$log = New-Object System.Text.StringBuilder
function L($m){ [void]$log.AppendLine($m); Set-Content -Path $logPath -Value $log.ToString() -Encoding UTF8 }
function Norm($p){ return ($p.Range.Text -replace "[`r`n`a]","").Trim() }

$path = "D:\이력서\ITT Cannon\Joonha_Lee_ITT_Cannon_FAE_Resume_Final_Integrated_v2.docx"
$word = New-Object -ComObject Word.Application
$word.Visible = $false
try { $word.DisplayAlerts = 0 } catch {}
$doc = $word.Documents.Open($path, $false, $false)
L('opened')

$HV='Alignment with ITT Purpose, Values & DNA'
$IV='Beyond technical fit, my working principles align with ITT''s purpose "We Solve It" and the three principles that define ITT''s high-performing culture and DNA:'
$BV1='Impeccable Character: 15+ years in defense and naval programs built on integrity, formal acceptance evidence (FAT/HAT/SAT), accurate documentation and customer trust - I commit to what the data supports and close issues with verifiable evidence.'
$BV2='Bold Thinking: I take on unmet technical challenges directly - capturing Voice of Customer, reframing field symptoms into root causes, and proposing practical design-in solutions that turn customer pain points into new design wins.'
$BV3='Collective Know-How: I act as the technical bridge across customer engineering, sales, R&D, product management, quality, production and suppliers - aligning people and facts to deliver solutions from concept through production.'
$BV4='"We Solve It" in practice: my evidence-based issue-closure method (symptom, reproduce, isolate the interface, corrective action, re-test, customer-facing evidence) is exactly how I would resolve ITT Cannon application and product issues with R&D support.'
$WEAVE='These working habits also map to ITT''s purpose, "We Solve It," and its guiding principles. My defense and naval background is grounded in Impeccable Character - integrity, accurate documentation and verifiable acceptance evidence. I apply Bold Thinking by taking on unmet technical challenges, capturing Voice of Customer and turning field symptoms into practical design-in solutions. And I rely on Collective Know-How, working as the technical bridge across customer engineering, sales, R&D, product management, quality and production to deliver results from concept through production.'

try {
  # 1) Weave Values substance into Self-Intro, before section 4 heading
  $anchor=$null
  foreach($p in $doc.Paragraphs){ if((Norm $p) -eq '4. Contribution After Joining ITT Cannon'){ $anchor=$p; break } }
  if($anchor -ne $null){
    $ins=$anchor.Range.Duplicate; $ins.Collapse(1)
    $ins.InsertBefore($WEAVE+"`r")
    L('woven values paragraph before section 4')
  } else { L('WARN: section 4 anchor not found') }

  # format the woven paragraph + bold key terms
  foreach($p in $doc.Paragraphs){
    if((Norm $p).StartsWith('These working habits also map to ITT')){
      $pr=$p.Range; $pr.Font.Name='Arial'; $pr.Font.Size=[single]10; $pr.Font.Bold=[int]0; $p.SpaceBefore=[single]0; $p.SpaceAfter=[single]3; $p.LineSpacingRule=0
      foreach($ph in @('We Solve It','Impeccable Character','Bold Thinking','Collective Know-How')){
        $ix=$p.Range.Text.IndexOf($ph)
        if($ix -ge 0){ $b=$p.Range.Duplicate; $b.Start=$p.Range.Start+$ix; $b.End=$p.Range.Start+$ix+$ph.Length; try{$b.Font.Bold=[int]1}catch{} }
      }
      break
    }
  }
  L('formatted woven paragraph')

  # 2) Remove standalone Values section (heading + intro + 4 bullets)
  $kill=@($HV,$IV,$BV1,$BV2,$BV3,$BV4)
  $targets=New-Object System.Collections.ArrayList
  foreach($p in $doc.Paragraphs){ if($kill -contains (Norm $p)){ [void]$targets.Add($p) } }
  $d=0
  for($k=$targets.Count-1;$k -ge 0;$k--){ try{ $r=$targets[$k].Range.Duplicate; $r.Delete()|Out-Null; $d++ }catch{ L('del fail') } }
  L("removed Values paragraphs=$d (expected 6)")

  # 3) Remove Requirement-Match table + its heading
  $reqTbl=$null
  foreach($t in $doc.Tables){ $fc=(($t.Cell(1,1).Range.Text) -replace "[`r`n`a]","").Trim(); if($fc -eq 'ITT Cannon FAE Requirement'){ $reqTbl=$t; break } }
  if($reqTbl -ne $null){ $reqTbl.Delete(); L('deleted requirement table') } else { L('WARN: requirement table not found') }
  $rh=0
  $hpars=New-Object System.Collections.ArrayList
  foreach($p in $doc.Paragraphs){ if((Norm $p) -eq 'Attachment: Condensed ITT Cannon FAE Requirement Match'){ [void]$hpars.Add($p) } }
  for($k=$hpars.Count-1;$k -ge 0;$k--){ try{ $r=$hpars[$k].Range.Duplicate; $r.Delete()|Out-Null; $rh++ }catch{} }
  L("removed requirement heading=$rh")

  # clean trailing empty paragraphs (collapse runs of blanks at end)
  $doc.Save(); L('saved')
  L('PAGES='+$doc.ComputeStatistics(2)+' TABLES='+$doc.Tables.Count)
  $vleft=$false; foreach($p in $doc.Paragraphs){ if((Norm $p) -eq $HV){$vleft=$true} }
  L("Values heading still present: $vleft")
}
catch { L('FATAL: '+$_.Exception.Message) }
finally {
  try{ $doc.Close([ref]$true) }catch{}
  try{ if($word.Documents.Count -eq 0){$word.Quit()} }catch{}
  try{ [Runtime.InteropServices.Marshal]::ReleaseComObject($word)|Out-Null }catch{}
}
Write-Output 'A_DONE'
