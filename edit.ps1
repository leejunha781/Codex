$ErrorActionPreference = 'Stop'
$log = New-Object System.Text.StringBuilder
function L($m){ [void]$log.AppendLine($m) }

$path = "D:\이력서\ITT Cannon\Joonha_Lee_ITT_Cannon_FAE_Resume_Standard_Format_v2.docx"
$pdf  = "D:\이력서\ITT Cannon\Joonha_Lee_ITT_Cannon_FAE_Resume_Standard_Format_v2.pdf"

$word = New-Object -ComObject Word.Application
$word.Visible = $false
try { $word.DisplayAlerts = 0 } catch {}
$doc = $word.Documents.Open($path, $false, $false)   # ConfirmConversions=F, ReadOnly=F

# ---------- helper: replace-all ----------
function ReplaceAll($find,$repl){
  $f = $doc.Content.Find
  $f.ClearFormatting()
  $f.Replacement.ClearFormatting()
  $r = $f.Execute($find,$true,$false,$false,$false,$false,$true,1,$false,$repl,2)
  return $r
}

# ---------- 1) Salary refinement (item #4) ----------
$r1 = ReplaceAll('KRW 90M+ base salary','KRW 90M+ base (above my current base of KRW 83.97M)')
$r2 = ReplaceAll('Minimum expectation is above my current basic salary.','Minimum expectation is a total package above my current total annual cash of approximately KRW 94.24M.')
L("SALARY replace1=$r1 replace2=$r2")

# ---------- 2) Insert Connector Product Knowledge block before 'Career Summary' (item #6) ----------
$HEADING = 'ITT Cannon Connector Product Knowledge & Target-Market Fit'
$INTRO   = 'ITT Cannon connector families mapped to the Field Application Engineer target markets, with my directly transferable experience:'
$B1 = 'Aerospace & Defense (preferred): MIL-DTL-38999 / D38999 circular (Cannon KJL / KJ / KJA-KJB), MKJ miniature circular, Micro-D and Nano-D high-density signal, and hermetic / filtered connectors - matched by 15+ years of ROK Navy ship and submarine communications, defense / aerospace electronics, and LEO satellite terminal work in harsh, EMI/EMC-critical environments.'
$B2 = 'Transportation & Industrial (Rail, Heavy Vehicles, Energy): APD and APV Vector rectangular, CA Bayonet (MIL-DTL-5015), high-power and mixed power-signal connectors, and EV (AC/DC) - matched by shipboard and industrial power / control cabling, board-to-system interconnects, production validation, and field installation experience.'
$B3 = 'Harsh-environment circular, electrical cabling, high-power / high-speed data (a plus): 38999 high-speed (KJA / KJAQ), high-speed data and fiber-optic (ARINC 801) lines, and high-power contacts - matched by RF / antenna and coaxial paths, Ethernet / TCP-UDP and high-speed data verification (Wireshark, iperf), and power / signal-integrity debugging.'
$B4 = 'Materials & manufacturing (high-performance thermoplastics, aluminum, copper alloys): composite (thermoplastic) 38999 shells, aluminum D-Sub, and copper-alloy contacts - matched by hands-on manufacturing and production-readiness background: ECO/BOM, pilot build, defective-part review, EMC/RoHS, and manufacturability assessment.'

# find Career Summary heading
$target = $null
foreach ($p in $doc.Paragraphs) {
  $t = ($p.Range.Text -replace "[`r`n`a]","").Trim()
  if ($t -eq 'Career Summary') { $target = $p; break }
}
if ($target -eq $null) { L('ERROR: Career Summary not found'); throw 'no anchor' }

$ins = $target.Range.Duplicate
$ins.Collapse(1)   # wdCollapseStart
$block = $HEADING + "`r" + $INTRO + "`r" + $B1 + "`r" + $B2 + "`r" + $B3 + "`r" + $B4 + "`r"
$ins.InsertBefore($block)
L('INSERTED product block')

# ---------- 3) Format the inserted block ----------
$bulletStyle = $doc.Styles.Item([int]-49)   # wdStyleListBullet
$bullets = @($B1,$B2,$B3,$B4)
foreach ($p in $doc.Paragraphs) {
  $t = ($p.Range.Text -replace "[`r`n`a]","").Trim()
  if ($t -eq $HEADING) {
     $pr = $p.Range
     $pr.Font.Name = 'Arial'; $pr.Font.Size = [single]14; $pr.Font.Bold = [int]1
     $p.SpaceBefore = [single]8; $p.SpaceAfter = [single]2; $p.LineSpacingRule = 0
  }
  elseif ($t -eq $INTRO) {
     $pr = $p.Range
     $pr.Font.Name = 'Arial'; $pr.Font.Size = [single]10; $pr.Font.Bold = [int]0
     $p.SpaceBefore = [single]0; $p.SpaceAfter = [single]2; $p.LineSpacingRule = 0
  }
  elseif ($bullets -contains $t) {
     $p.Style = $bulletStyle
     $pr = $p.Range
     $pr.Font.Name = 'Arial'; $pr.Font.Size = [single]10; $pr.Font.Bold = [int]0
     $p.SpaceBefore = [single]0; $p.SpaceAfter = [single]2; $p.LineSpacingRule = 0
     # bold the lead-in label up to first colon
     $idx = $p.Range.Text.IndexOf(':')
     if ($idx -gt 0) {
        $lbl = $p.Range.Duplicate
        $lbl.End = $p.Range.Start + $idx + 1
        $lbl.Font.Bold = [int]1
     }
  }
}
L('FORMATTED product block')

# ---------- 4) Normalize 'Compensation / Salary Information' heading ----------
$normal = $doc.Styles.Item([int]-1)  # wdStyleNormal
foreach ($p in $doc.Paragraphs) {
  $t = ($p.Range.Text -replace "[`r`n`a]","").Trim()
  if ($t -eq 'Compensation / Salary Information') {
     $p.Style = $normal
     $pr = $p.Range
     $pr.Font.Name = 'Arial'; $pr.Font.Size = [single]14; $pr.Font.Bold = [int]1
     $p.SpaceBefore = [single]8; $p.SpaceAfter = [single]2; $p.LineSpacingRule = 0
     L('NORMALIZED compensation heading')
  }
}

# ---------- 5) Global pass: Arial, single spacing, bump sizes < 10 to 10 ----------
$bumped = 0
foreach ($p in $doc.Paragraphs) {
  $pr = $p.Range
  try { $pr.Font.Name = 'Arial' } catch {}
  try { $p.LineSpacingRule = 0 } catch {}
  $sz = $null
  try { $sz = $pr.Font.Size } catch {}
  if ($sz -ne $null) {
    if ($sz -eq 9999999) {
      foreach ($w in $pr.Words) {
        $ws = $null; try { $ws = $w.Font.Size } catch {}
        if ($ws -ne $null -and $ws -ne 9999999 -and $ws -lt 10) { $w.Font.Size = [single]10; $bumped++ }
      }
    } elseif ($sz -lt 10) {
      $pr.Font.Size = [single]10; $bumped++
    }
  }
}
L("GLOBAL bumped paragraphs/words to 10pt: $bumped")

# ---------- 6) Subtle bottom border on section headings (size 14 + bold) ----------
$bordered = 0
foreach ($p in $doc.Paragraphs) {
  $pr = $p.Range
  $sz = $null; try { $sz = $pr.Font.Size } catch {}
  $bd = $null; try { $bd = $pr.Font.Bold } catch {}
  if ($sz -eq 14 -and $bd -eq -1) {
     try {
        $bb = $pr.ParagraphFormat.Borders.Item([int]-3)  # wdBorderBottom
        $bb.LineStyle = [int]1     # wdLineStyleSingle
        $bb.LineWidth = [int]6     # 0.75 pt
        $bb.Color = [int]5855577   # gray (89,89,89) BGR
        $bordered++
     } catch { L('border fail: ' + $_.Exception.Message) }
  }
}
L("HEADINGS bordered: $bordered")

# ---------- 7) Save + export PDF ----------
$doc.Save()
L('SAVED docx')
try { $doc.ExportAsFixedFormat($pdf, 17); L('EXPORTED pdf') } catch { L('PDF export fail: ' + $_.Exception.Message) }

# ---------- 8) QA report ----------
L('----- QA -----')
L('PAGES=' + $doc.ComputeStatistics(2))
L('WORDS=' + $doc.ComputeStatistics(0))
L('TABLES=' + $doc.Tables.Count)
$min = 9999
foreach ($p in $doc.Paragraphs) {
  $sz = $null; try { $sz = $p.Range.Font.Size } catch {}
  if ($sz -ne $null -and $sz -ne 9999999 -and $sz -lt $min) { $min = $sz }
}
L("MIN_FONT_SIZE=$min")
L('--- section headings (size14 bold) ---')
foreach ($p in $doc.Paragraphs) {
  $sz=$null; try {$sz=$p.Range.Font.Size} catch{}
  $bd=$null; try {$bd=$p.Range.Font.Bold} catch{}
  if ($sz -eq 14 -and $bd -eq -1) { L('  H: ' + (($p.Range.Text -replace "[`r`n`a]","").Trim())) }
}
L('--- inserted block readback ---')
$want = @($HEADING,$INTRO,$B1,$B2,$B3,$B4)
foreach ($p in $doc.Paragraphs) {
  $t = ($p.Range.Text -replace "[`r`n`a]","").Trim()
  if ($want -contains $t) {
     $st=''; try {$st=$p.Style.NameLocal} catch{}
     L('  OK[' + $p.Range.Font.Size + '/' + $st + ']: ' + $t.Substring(0,[Math]::Min(55,$t.Length)))
  }
}

Set-Content -Path "C:\Users\namma\.claude\itt_work\edit_log.txt" -Value $log.ToString() -Encoding UTF8

$doc.Close([ref]$true)
if ($word.Documents.Count -eq 0) { $word.Quit() }
[Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
Write-Output 'EDIT_DONE'
