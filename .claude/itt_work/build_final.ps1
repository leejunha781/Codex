$ErrorActionPreference = 'Continue'
$logPath = "C:\Users\namma\.claude\itt_work\build_log.txt"
$log = New-Object System.Text.StringBuilder
function L($m){ [void]$log.AppendLine($m); Set-Content -Path $logPath -Value $log.ToString() -Encoding UTF8 }
function Norm($p){ return ($p.Range.Text -replace "[`r`n`a]","").Trim() }

$path = "D:\이력서\ITT Cannon\Joonha_Lee_ITT_Cannon_FAE_Resume_Final_Integrated_v2.docx"
$pdf  = "D:\이력서\ITT Cannon\Joonha_Lee_ITT_Cannon_FAE_Resume_Final_Integrated_v2.pdf"

$word = New-Object -ComObject Word.Application
$word.Visible = $false
try { $word.DisplayAlerts = 0 } catch {}
L('OPENED Word')
$doc = $word.Documents.Open($path, $false, $false)
L('OPENED doc')

try {
  # ---------- 1) Salary cells ----------
  $r1=$false;$r2=$false
  foreach ($tbl in $doc.Tables){ foreach ($c in $tbl.Range.Cells){
    $ct=($c.Range.Text -replace "[`r`n`a]","").Trim()
    if ($ct -eq 'KRW 90M+ base salary'){ $c.Range.Text='Above my current base salary (KRW 83.97M); negotiable'; $r1=$true }
    elseif ($ct.StartsWith('Negotiable depending on total compensation')){ $c.Range.Text='Open and negotiable based on total compensation, role level and incentive structure within the ITT Cannon FAE market range. Minimum expectation is a total package above my current total annual cash of approximately KRW 94.24M.'; $r2=$true }
  } }
  L("STEP1 salary r1=$r1 r2=$r2")

  # ---------- 2) Split wall-of-text body paragraphs (style Normal, not table, size10, not bold, len>340) ----------
  function SplitPara($p,$budget){
    $txt=($p.Range.Text -replace "[`r`n`a]","").Trim()
    $parts=[System.Text.RegularExpressions.Regex]::Split($txt,'(?<=\.)\s+(?=[A-Z(])')
    $chunks=@(); $cur=''
    foreach($s in $parts){ if($cur -eq ''){$cur=$s} elseif(($cur.Length+1+$s.Length) -le $budget){$cur=$cur+' '+$s} else {$chunks+=$cur; $cur=$s} }
    if($cur -ne ''){$chunks+=$cur}
    if($chunks.Count -lt 2){ return $false }
    $new=($chunks -join "`r")
    $r=$p.Range.Duplicate; $r.End=$r.End-1; $r.Text=$new
    return $true
  }
  $targets=New-Object System.Collections.ArrayList
  foreach($p in $doc.Paragraphs){
    $intbl=$false; try{$intbl=$p.Range.Information(12)}catch{}
    if($intbl){continue}
    $st='';try{$st=$p.Style.NameLocal}catch{}
    if($st -ne '표준'){continue}
    $sz=$null;try{$sz=$p.Range.Font.Size}catch{}
    $bd=$null;try{$bd=$p.Range.Font.Bold}catch{}
    $len=(Norm $p).Length
    if($sz -eq 10 -and $bd -eq 0 -and $len -gt 340){ [void]$targets.Add($p) }
  }
  $sc=0
  for($k=$targets.Count-1;$k -ge 0;$k--){ if(SplitPara $targets[$k] 320){$sc++} }
  L("STEP2 split paragraphs=$sc (targets=$($targets.Count))")

  # ---------- 3) Insert two new sections ----------
  $HC='ITT Cannon Connector Product Knowledge & Target-Market Fit'
  $IC='ITT Cannon connector families mapped to the Field Application Engineer target markets, with my directly transferable experience:'
  $BC1='Aerospace & Defense (preferred): MIL-DTL-38999 / D38999 circular (Cannon KJL / KJ / KJA-KJB), MKJ miniature circular, Micro-D and Nano-D high-density signal, and hermetic / filtered connectors - matched by 15+ years of ROK Navy ship and submarine communications, defense / aerospace electronics, and LEO satellite terminal work in harsh, EMI/EMC-critical environments.'
  $BC2='Transportation & Industrial (Rail, Heavy Vehicles, Energy): APD and APV Vector rectangular, CA Bayonet (MIL-DTL-5015), high-power and mixed power-signal connectors, and EV (AC/DC) - matched by shipboard and industrial power / control cabling, board-to-system interconnects, production validation, and field installation experience.'
  $BC3='Harsh-environment circular, electrical cabling, high-power / high-speed data (a plus): 38999 high-speed (KJA / KJAQ), high-speed data and fiber-optic (ARINC 801) lines, and high-power contacts - matched by RF / antenna and coaxial paths, Ethernet / TCP-UDP and high-speed data verification (Wireshark, iperf), and power / signal-integrity debugging.'
  $BC4='Materials & manufacturing (high-performance thermoplastics, aluminum, copper alloys): composite (thermoplastic) 38999 shells, aluminum D-Sub, and copper-alloy contacts - matched by hands-on manufacturing and production-readiness background: ECO/BOM, pilot build, defective-part review, EMC/RoHS, and manufacturability assessment.'

  $HV='Alignment with ITT Purpose, Values & DNA'
  $IV='Beyond technical fit, my working principles align with ITT''s purpose "We Solve It" and the three principles that define ITT''s high-performing culture and DNA:'
  $BV1='Impeccable Character: 15+ years in defense and naval programs built on integrity, formal acceptance evidence (FAT/HAT/SAT), accurate documentation and customer trust - I commit to what the data supports and close issues with verifiable evidence.'
  $BV2='Bold Thinking: I take on unmet technical challenges directly - capturing Voice of Customer, reframing field symptoms into root causes, and proposing practical design-in solutions that turn customer pain points into new design wins.'
  $BV3='Collective Know-How: I act as the technical bridge across customer engineering, sales, R&D, product management, quality, production and suppliers - aligning people and facts to deliver solutions from concept through production.'
  $BV4='"We Solve It" in practice: my evidence-based issue-closure method (symptom, reproduce, isolate the interface, corrective action, re-test, customer-facing evidence) is exactly how I would resolve ITT Cannon application and product issues with R&D support.'

  function InsertBlockBefore($anchorText,$lines){
    $target=$null
    foreach($p in $doc.Paragraphs){ if((Norm $p) -eq $anchorText){ $target=$p; break } }
    if($target -eq $null){ L("anchor NOT found: $anchorText"); return $false }
    $ins=$target.Range.Duplicate; $ins.Collapse(1)
    $block=($lines -join "`r")+"`r"
    $ins.InsertBefore($block)
    return $true
  }
  $okC=InsertBlockBefore 'Education, Certifications, Language and Technical Toolkit' @($HC,$IC,$BC1,$BC2,$BC3,$BC4)
  $okV=InsertBlockBefore 'Self-Introduction and Application Motivation' @($HV,$IV,$BV1,$BV2,$BV3,$BV4)
  L("STEP3 inserts connector=$okC values=$okV")

  # ---------- 4) Format inserted blocks ----------
  $bulletStyle=$null; try{$bulletStyle=$doc.Styles.Item([int]-49)}catch{L('bullet style fail')}
  $headings=@($HC,$HV); $intros=@($IC,$IV); $bullets=@($BC1,$BC2,$BC3,$BC4,$BV1,$BV2,$BV3,$BV4)
  foreach($p in $doc.Paragraphs){
    $t=Norm $p
    if($headings -contains $t){
      $pr=$p.Range; $pr.Font.Name='Arial'; $pr.Font.Size=[single]14; $pr.Font.Bold=[int]1
      $p.SpaceBefore=[single]8; $p.SpaceAfter=[single]3; $p.LineSpacingRule=0
    } elseif($intros -contains $t){
      $pr=$p.Range; $pr.Font.Name='Arial'; $pr.Font.Size=[single]10; $pr.Font.Bold=[int]0
      $p.SpaceBefore=[single]0; $p.SpaceAfter=[single]3; $p.LineSpacingRule=0
    } elseif($bullets -contains $t){
      if($bulletStyle -ne $null){ try{$p.Style=$bulletStyle}catch{} }
      $pr=$p.Range; $pr.Font.Name='Arial'; $pr.Font.Size=[single]10; $pr.Font.Bold=[int]0
      $p.SpaceBefore=[single]0; $p.SpaceAfter=[single]3; $p.LineSpacingRule=0
      $idx=$p.Range.Text.IndexOf(':')
      if($idx -gt 0){ $lbl=$p.Range.Duplicate; $lbl.End=$p.Range.Start+$idx+1; $lbl.Font.Bold=[int]1 }
    }
  }
  L('STEP4 formatted inserted blocks')

  # ---------- 5) Bold JD-key phrases in body prose (style Normal, not table) ----------
  $phrases=@('21+ years','Voice of Customer','new design wins','evidence-based problem solving','harsh-environment connectors','cross-functional coordination')
  $done=@{}
  foreach($p in $doc.Paragraphs){
    $intbl=$false; try{$intbl=$p.Range.Information(12)}catch{}
    if($intbl){continue}
    $st='';try{$st=$p.Style.NameLocal}catch{}
    if($st -ne '표준'){continue}
    $ptxt=$p.Range.Text
    foreach($ph in $phrases){
      if(-not $done.ContainsKey($ph)){
        $ix=$ptxt.IndexOf($ph)
        if($ix -ge 0){ $b=$p.Range.Duplicate; $b.Start=$p.Range.Start+$ix; $b.End=$p.Range.Start+$ix+$ph.Length; try{$b.Font.Bold=[int]1; $done[$ph]=$true}catch{} }
      }
    }
  }
  L("STEP5 bolded phrases=$($done.Count)")

  # ---------- 6) Global: Arial, single spacing, bump <10 -> 10 ----------
  $bumped=0
  foreach($p in $doc.Paragraphs){
    $pr=$p.Range
    try{$pr.Font.Name='Arial'}catch{}
    try{$p.LineSpacingRule=0}catch{}
    $sz=$null;try{$sz=$pr.Font.Size}catch{}
    if($sz -ne $null){
      if($sz -eq 9999999){ foreach($w in $pr.Words){ $ws=$null;try{$ws=$w.Font.Size}catch{}; if($ws -ne $null -and $ws -ne 9999999 -and $ws -lt 10){$w.Font.Size=[single]10;$bumped++} } }
      elseif($sz -lt 10){ $pr.Font.Size=[single]10; $bumped++ }
    }
  }
  L("STEP6 bumped=$bumped")

  # ---------- 7) Stronger heading borders (navy, 1.5pt) ----------
  # detect heading color from an existing section heading
  $hcolor=8210719  # default navy RGB(31,73,125) BGR
  foreach($p in $doc.Paragraphs){ if((Norm $p) -eq 'Professional Profile and Interest in ITT Cannon FAE Role'){ try{ $cc=$p.Range.Font.Color; if($cc -gt 0 -and $cc -ne 9999999){$hcolor=$cc} }catch{}; break } }
  L("STEP7 heading color=$hcolor")
  $bordered=0
  foreach($p in $doc.Paragraphs){
    $pr=$p.Range
    $sz=$null;try{$sz=$pr.Font.Size}catch{}
    $bd=$null;try{$bd=$pr.Font.Bold}catch{}
    if($sz -eq 14 -and $bd -eq -1){
      try{ $bb=$pr.ParagraphFormat.Borders.Item([int]-3); $bb.LineStyle=[int]1; $bb.LineWidth=[int]12; $bb.Color=[int]$hcolor; $bordered++ }catch{ L('border fail: '+$_.Exception.Message) }
    }
  }
  L("STEP7 bordered=$bordered")

  # ---------- 8) Save + PDF ----------
  $doc.Save(); L('STEP8 saved')
  try{ $doc.ExportAsFixedFormat($pdf,17); L('STEP8 pdf') }catch{ L('pdf fail: '+$_.Exception.Message) }

  # ---------- 9) QA ----------
  L('PAGES='+$doc.ComputeStatistics(2)+' WORDS='+$doc.ComputeStatistics(0)+' TABLES='+$doc.Tables.Count)
  $min=9999; foreach($p in $doc.Paragraphs){ $sz=$null;try{$sz=$p.Range.Font.Size}catch{}; if($sz -ne $null -and $sz -ne 9999999 -and $sz -lt $min){$min=$sz} }
  L("MIN_FONT=$min")
  L('--- headings(14/bold) ---')
  foreach($p in $doc.Paragraphs){ $sz=$null;try{$sz=$p.Range.Font.Size}catch{}; $bd=$null;try{$bd=$p.Range.Font.Bold}catch{}; if($sz -eq 14 -and $bd -eq -1){ L('  H: '+(Norm $p)) } }
  L('--- inserted readback ---')
  $want=@($HC,$HV,$BC1,$BV1,$BV4)
  foreach($p in $doc.Paragraphs){ $t=Norm $p; if($want -contains $t){ $st='';try{$st=$p.Style.NameLocal}catch{}; L('  OK[sz='+$p.Range.Font.Size+' '+$st+']: '+$t.Substring(0,[Math]::Min(46,$t.Length))) } }
  L('--- salary readback ---')
  foreach($tbl in $doc.Tables){ foreach($c in $tbl.Range.Cells){ $ct=($c.Range.Text -replace "[`r`n`a]","").Trim(); if($ct.StartsWith('Above my current base') -or $ct.StartsWith('Open and negotiable')){ L('  $: '+$ct.Substring(0,[Math]::Min(60,$ct.Length))) } } }
}
catch { L('FATAL: '+$_.Exception.Message) }
finally {
  try{ $doc.Close([ref]$true) }catch{ L('close fail: '+$_.Exception.Message) }
  try{ if($word.Documents.Count -eq 0){$word.Quit()} }catch{}
  try{ [Runtime.InteropServices.Marshal]::ReleaseComObject($word)|Out-Null }catch{}
}
Write-Output 'BUILD_FINAL_DONE'
