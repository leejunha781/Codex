$ErrorActionPreference = 'Continue'
$logPath = "C:\Users\namma\.claude\itt_work\edit_log.txt"
$log = New-Object System.Text.StringBuilder
function L($m){ [void]$log.AppendLine($m); Set-Content -Path $logPath -Value $log.ToString() -Encoding UTF8 }

$path = "D:\이력서\ITT Cannon\Joonha_Lee_ITT_Cannon_FAE_Resume_Standard_Format_v2.docx"
$pdf  = "D:\이력서\ITT Cannon\Joonha_Lee_ITT_Cannon_FAE_Resume_Standard_Format_v2.pdf"

$word = New-Object -ComObject Word.Application
$word.Visible = $false
try { $word.DisplayAlerts = 0 } catch {}
L('OPENED Word instance')
$doc = $word.Documents.Open($path, $false, $false)
L('OPENED doc')

try {
  # 1) Salary refinement via direct cell text (no Find.Execute)
  $r1=$false; $r2=$false
  foreach ($tbl in $doc.Tables) {
    foreach ($c in $tbl.Range.Cells) {
      $ct = ($c.Range.Text -replace "[`r`n`a]","").Trim()
      if ($ct -eq 'KRW 90M+ base salary') {
        $c.Range.Text = 'KRW 90M+ base (above my current base of KRW 83.97M)'; $r1=$true
      } elseif ($ct.StartsWith('Negotiable depending on total compensation')) {
        $c.Range.Text = 'Negotiable on base, incentive and role level to the ITT Cannon FAE market range. Minimum expectation is a total package above my current total annual cash of approximately KRW 94.24M.'; $r2=$true
      }
    }
  }
  L("STEP1 salary r1=$r1 r2=$r2")

  # 2) Insert product block before 'Career Summary'
  $HEADING = 'ITT Cannon Connector Product Knowledge & Target-Market Fit'
  $INTRO   = 'ITT Cannon connector families mapped to the Field Application Engineer target markets, with my directly transferable experience:'
  $B1 = 'Aerospace & Defense (preferred): MIL-DTL-38999 / D38999 circular (Cannon KJL / KJ / KJA-KJB), MKJ miniature circular, Micro-D and Nano-D high-density signal, and hermetic / filtered connectors - matched by 15+ years of ROK Navy ship and submarine communications, defense / aerospace electronics, and LEO satellite terminal work in harsh, EMI/EMC-critical environments.'
  $B2 = 'Transportation & Industrial (Rail, Heavy Vehicles, Energy): APD and APV Vector rectangular, CA Bayonet (MIL-DTL-5015), high-power and mixed power-signal connectors, and EV (AC/DC) - matched by shipboard and industrial power / control cabling, board-to-system interconnects, production validation, and field installation experience.'
  $B3 = 'Harsh-environment circular, electrical cabling, high-power / high-speed data (a plus): 38999 high-speed (KJA / KJAQ), high-speed data and fiber-optic (ARINC 801) lines, and high-power contacts - matched by RF / antenna and coaxial paths, Ethernet / TCP-UDP and high-speed data verification (Wireshark, iperf), and power / signal-integrity debugging.'
  $B4 = 'Materials & manufacturing (high-performance thermoplastics, aluminum, copper alloys): composite (thermoplastic) 38999 shells, aluminum D-Sub, and copper-alloy contacts - matched by hands-on manufacturing and production-readiness background: ECO/BOM, pilot build, defective-part review, EMC/RoHS, and manufacturability assessment.'

  $target=$null
  foreach ($p in $doc.Paragraphs) { $t=($p.Range.Text -replace "[`r`n`a]","").Trim(); if ($t -eq 'Career Summary'){ $target=$p; break } }
  if ($target -eq $null) { throw 'Career Summary anchor not found' }
  $ins=$target.Range.Duplicate; $ins.Collapse(1)
  $block = $HEADING+"`r"+$INTRO+"`r"+$B1+"`r"+$B2+"`r"+$B3+"`r"+$B4+"`r"
  $ins.InsertBefore($block)
  L('STEP2 inserted block')

  # 3) Format inserted block
  $bulletStyle=$null; try { $bulletStyle=$doc.Styles.Item([int]-49) } catch { L('bullet style lookup failed') }
  $bullets=@($B1,$B2,$B3,$B4)
  foreach ($p in $doc.Paragraphs) {
    $t=($p.Range.Text -replace "[`r`n`a]","").Trim()
    if ($t -eq $HEADING) {
      $pr=$p.Range; $pr.Font.Name='Arial'; $pr.Font.Size=[single]14; $pr.Font.Bold=[int]1
      $p.SpaceBefore=[single]8; $p.SpaceAfter=[single]2; $p.LineSpacingRule=0
    } elseif ($t -eq $INTRO) {
      $pr=$p.Range; $pr.Font.Name='Arial'; $pr.Font.Size=[single]10; $pr.Font.Bold=[int]0
      $p.SpaceBefore=[single]0; $p.SpaceAfter=[single]2; $p.LineSpacingRule=0
    } elseif ($bullets -contains $t) {
      if ($bulletStyle -ne $null) { try { $p.Style=$bulletStyle } catch { L('apply bullet style failed') } }
      $pr=$p.Range; $pr.Font.Name='Arial'; $pr.Font.Size=[single]10; $pr.Font.Bold=[int]0
      $p.SpaceBefore=[single]0; $p.SpaceAfter=[single]2; $p.LineSpacingRule=0
      $idx=$p.Range.Text.IndexOf(':')
      if ($idx -gt 0) { $lbl=$p.Range.Duplicate; $lbl.End=$p.Range.Start+$idx+1; $lbl.Font.Bold=[int]1 }
    }
  }
  L('STEP3 formatted block')

  # 4) Normalize Compensation heading
  $normal=$null; try { $normal=$doc.Styles.Item([int]-1) } catch {}
  foreach ($p in $doc.Paragraphs) {
    $t=($p.Range.Text -replace "[`r`n`a]","").Trim()
    if ($t -eq 'Compensation / Salary Information') {
      if ($normal -ne $null) { try { $p.Style=$normal } catch {} }
      $pr=$p.Range; $pr.Font.Name='Arial'; $pr.Font.Size=[single]14; $pr.Font.Bold=[int]1
      $p.SpaceBefore=[single]8; $p.SpaceAfter=[single]2; $p.LineSpacingRule=0
      L('STEP4 normalized compensation heading')
    }
  }

  # 5) Global Arial + single + bump <10 -> 10
  $bumped=0
  foreach ($p in $doc.Paragraphs) {
    $pr=$p.Range
    try { $pr.Font.Name='Arial' } catch {}
    try { $p.LineSpacingRule=0 } catch {}
    $sz=$null; try { $sz=$pr.Font.Size } catch {}
    if ($sz -ne $null) {
      if ($sz -eq 9999999) {
        foreach ($w in $pr.Words){ $ws=$null; try{$ws=$w.Font.Size}catch{}; if ($ws -ne $null -and $ws -ne 9999999 -and $ws -lt 10){ $w.Font.Size=[single]10; $bumped++ } }
      } elseif ($sz -lt 10) { $pr.Font.Size=[single]10; $bumped++ }
    }
  }
  L("STEP5 global bumped=$bumped")

  # 6) Heading bottom borders (size 14 + bold)
  $bordered=0
  foreach ($p in $doc.Paragraphs) {
    $pr=$p.Range
    $sz=$null; try{$sz=$pr.Font.Size}catch{}
    $bd=$null; try{$bd=$pr.Font.Bold}catch{}
    if ($sz -eq 14 -and $bd -eq -1) {
      try { $bb=$pr.ParagraphFormat.Borders.Item([int]-3); $bb.LineStyle=[int]1; $bb.LineWidth=[int]6; $bb.Color=[int]5855577; $bordered++ } catch { L('border fail: '+$_.Exception.Message) }
    }
  }
  L("STEP6 bordered=$bordered")

  # 7) Save + PDF
  $doc.Save(); L('STEP7 saved')
  try { $doc.ExportAsFixedFormat($pdf,17); L('STEP7 pdf exported') } catch { L('pdf fail: '+$_.Exception.Message) }

  # 8) QA
  L('PAGES='+$doc.ComputeStatistics(2)); L('WORDS='+$doc.ComputeStatistics(0)); L('TABLES='+$doc.Tables.Count)
  $min=9999; foreach ($p in $doc.Paragraphs){ $sz=$null;try{$sz=$p.Range.Font.Size}catch{}; if ($sz -ne $null -and $sz -ne 9999999 -and $sz -lt $min){$min=$sz} }
  L("MIN_FONT_SIZE=$min")
  L('--- headings (14/bold) ---')
  foreach ($p in $doc.Paragraphs){ $sz=$null;try{$sz=$p.Range.Font.Size}catch{}; $bd=$null;try{$bd=$p.Range.Font.Bold}catch{}; if ($sz -eq 14 -and $bd -eq -1){ L('  H: '+(($p.Range.Text -replace "[`r`n`a]","").Trim())) } }
  L('--- inserted readback ---')
  $want=@($HEADING,$INTRO,$B1,$B2,$B3,$B4)
  foreach ($p in $doc.Paragraphs){ $t=($p.Range.Text -replace "[`r`n`a]","").Trim(); if ($want -contains $t){ $st='';try{$st=$p.Style.NameLocal}catch{}; L('  OK[sz='+$p.Range.Font.Size+' style='+$st+']: '+$t.Substring(0,[Math]::Min(48,$t.Length))) } }
  L('--- expected comp readback ---')
  foreach ($tbl in $doc.Tables){ foreach ($c in $tbl.Range.Cells){ $ct=($c.Range.Text -replace "[`r`n`a]","").Trim(); if ($ct.StartsWith('KRW 90M+ base (') -or $ct.StartsWith('Negotiable on base')){ L('  COMP: '+$ct.Substring(0,[Math]::Min(60,$ct.Length))) } } }
}
catch { L('FATAL: ' + $_.Exception.Message) }
finally {
  try { $doc.Close([ref]$true) } catch { L('close fail: '+$_.Exception.Message) }
  try { if ($word.Documents.Count -eq 0) { $word.Quit() } } catch {}
  try { [Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null } catch {}
}
Write-Output 'EDIT3_DONE'
