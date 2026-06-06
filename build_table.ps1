$ErrorActionPreference = 'Continue'
$logPath = "C:\Users\namma\.claude\itt_work\table_log.txt"
$log = New-Object System.Text.StringBuilder
function L($m){ [void]$log.AppendLine($m); Set-Content -Path $logPath -Value $log.ToString() -Encoding UTF8 }
function Norm($p){ return ($p.Range.Text -replace "[`r`n`a]","").Trim() }

$path = "D:\이력서\ITT Cannon\Joonha_Lee_ITT_Cannon_FAE_Resume_Final_Integrated_v2.docx"
$pdf  = "D:\이력서\ITT Cannon\Joonha_Lee_ITT_Cannon_FAE_Resume_Final_Integrated_v2.pdf"

# Border style chosen by user: charcoal gray 1.5pt
$bWidth = 12          # 1.5pt (eighths of a point)
$bColor = 4210752     # RGB(64,64,64) charcoal, BGR

$word = New-Object -ComObject Word.Application
$word.Visible = $false
try { $word.DisplayAlerts = 0 } catch {}
$doc = $word.Documents.Open($path, $false, $false)
L('opened')

try {
  # detect existing table-header fill (from Career Summary 'Period' cell) for consistency
  $tblHdr = 7949855
  foreach($t in $doc.Tables){ foreach($c in $t.Range.Cells){ if((($c.Range.Text) -replace "[`r`n`a]","").Trim() -eq 'Period'){ try{$v=$c.Shading.BackgroundPatternColor; if($v -ne -16777216 -and $v -ge 0){$tblHdr=$v}}catch{}; break } } }
  L("tblHdr=$tblHdr")

  # ---- 1) Append heading + requirement-match table at end ----
  $cnt=$doc.Paragraphs.Count
  $last=$doc.Paragraphs.Item($cnt)
  $ir=$last.Range.Duplicate; $ir.Collapse(1)
  $ir.InsertBefore("Attachment: Condensed ITT Cannon FAE Requirement Match`r")
  # format heading
  foreach($p in $doc.Paragraphs){ if((Norm $p) -eq 'Attachment: Condensed ITT Cannon FAE Requirement Match'){ $pr=$p.Range; $pr.Font.Name='Arial'; $pr.Font.Size=[single]14; $pr.Font.Bold=[int]1; $p.SpaceBefore=[single]8; $p.SpaceAfter=[single]3; $p.LineSpacingRule=0 } }

  # add table before final empty paragraph
  $cnt2=$doc.Paragraphs.Count
  $last2=$doc.Paragraphs.Item($cnt2)
  $tr=$last2.Range.Duplicate; $tr.Collapse(1)
  $tbl=$doc.Tables.Add($tr,8,3)
  L('table added')

  $data=@(
    @('ITT Cannon FAE Requirement','Condensed Career Evidence','Value to ITT Cannon'),
    @('Recommend products and technical solutions for new designs','Converted customer requirements into architecture, interface requirements, technical proposals, test plans and production criteria.','Supports design-in by linking application conditions to connector and cabling solutions.'),
    @('Identify design requirements and unmet challenges; capture Voice of Customer','Engaged Navy users, shipyards, inspectors, Eutelsat OneWeb, Marlink/STW, suppliers and internal engineering teams.','Captures VOC and translates customer pain points into practical application requirements.'),
    @('Primary contact between customer engineering and internal cross-functional teams','Coordinated engineering, quality, production, procurement, suppliers, subcontractors and customers through installation, FAT/HAT/SAT, commissioning and issue closure.','Acts as a reliable bridge among customer engineering, sales, R&D, product management and operations.'),
    @('Solve application / product issues with R&D support','Used board debugging, RF measurements, logs, defect history, packet analysis, environmental data and re-test evidence.','Escalates issues with clear reproduction data, evidence and corrective-action tracking.'),
    @('Develop and present technical solutions to capture design wins','Prepared technical proposals, WBS schedules, cost estimates, test procedures, result reports and customer-facing summaries.','Presents ITT Cannon value clearly and supports new design opportunities.'),
    @('Preferred Aerospace & Defense market experience in Korea','15+ years ROK Navy ship / submarine systems, plus defense / aerospace electronics and LEO satellite terminals.','Strong fit for Korean A&D customers requiring reliability and harsh-environment performance.'),
    @('Harsh-environment connectors, cabling, high-power / high-speed data; materials & manufacturing','Shipboard electrical cabling, RF antenna/cable interfaces, Ethernet/high-speed data, power/control/signal interfaces; thermoplastic/aluminum/copper and production background.','Understands rugged interconnect use cases, power/signal integrity, RF/data transfer, materials and field-installation limits.')
  )
  for($r=0;$r -lt 8;$r++){ for($c=0;$c -lt 3;$c++){ $tbl.Cell($r+1,$c+1).Range.Text=$data[$r][$c] } }
  L('table filled')

  # table formatting
  $tbl.Range.Font.Name='Arial'; $tbl.Range.Font.Size=[single]10
  $tbl.Range.ParagraphFormat.SpaceAfter=[single]2; $tbl.Range.ParagraphFormat.SpaceBefore=[single]2; $tbl.Range.ParagraphFormat.LineSpacingRule=0
  try{ $tbl.Borders.InsideLineStyle=[int]1; $tbl.Borders.OutsideLineStyle=[int]1 }catch{ L('tbl border fail') }
  $tbl.PreferredWidthType=2; $tbl.PreferredWidth=[single]100
  try{ $tbl.Columns.Item(1).PreferredWidthType=2; $tbl.Columns.Item(1).PreferredWidth=[single]30 }catch{}
  try{ $tbl.Columns.Item(2).PreferredWidthType=2; $tbl.Columns.Item(2).PreferredWidth=[single]42 }catch{}
  try{ $tbl.Columns.Item(3).PreferredWidthType=2; $tbl.Columns.Item(3).PreferredWidth=[single]28 }catch{}
  # header row styling
  for($c=1;$c -le 3;$c++){ $cell=$tbl.Cell(1,$c); try{$cell.Shading.BackgroundPatternColor=[int]$tblHdr}catch{}; $cell.Range.Font.Bold=[int]1; try{$cell.Range.Font.Color=[int]16777215}catch{} }
  L('table styled')

  # ---- 2) Re-apply ALL heading underlines as charcoal 1.5pt (non-empty 14/bold) ----
  $bordered=0
  foreach($p in $doc.Paragraphs){
    if((Norm $p).Length -eq 0){ continue }
    $pr=$p.Range
    $sz=$null;try{$sz=$pr.Font.Size}catch{}
    $bd=$null;try{$bd=$pr.Font.Bold}catch{}
    if($sz -eq 14 -and $bd -eq -1){
      try{ $bb=$pr.ParagraphFormat.Borders.Item([int]-3); $bb.LineStyle=[int]1; $bb.LineWidth=[int]$bWidth; $bb.Color=[int]$bColor; $bordered++ }catch{ L('border fail: '+$_.Exception.Message) }
    }
  }
  L("re-bordered=$bordered (charcoal $bWidth/8 pt)")

  # safety: ensure no sub-10 text introduced
  $bumped=0
  foreach($p in $doc.Paragraphs){ $pr=$p.Range; $sz=$null;try{$sz=$pr.Font.Size}catch{}; if($sz -ne $null -and $sz -ne 9999999 -and $sz -lt 10){ $pr.Font.Size=[single]10; $bumped++ } }
  L("bumped=$bumped")

  $doc.Save(); L('saved')
  try{ $doc.ExportAsFixedFormat($pdf,17); L('pdf') }catch{ L('pdf fail: '+$_.Exception.Message) }

  L('PAGES='+$doc.ComputeStatistics(2)+' TABLES='+$doc.Tables.Count)
  $min=9999; foreach($p in $doc.Paragraphs){ $sz=$null;try{$sz=$p.Range.Font.Size}catch{}; if($sz -ne $null -and $sz -ne 9999999 -and $sz -lt $min){$min=$sz} }
  L("MIN_FONT=$min")
  L('--- headings ---')
  foreach($p in $doc.Paragraphs){ $sz=$null;try{$sz=$p.Range.Font.Size}catch{}; $bd=$null;try{$bd=$p.Range.Font.Bold}catch{}; if($sz -eq 14 -and $bd -eq -1 -and (Norm $p).Length -gt 0){ L('  H: '+(Norm $p)) } }
}
catch { L('FATAL: '+$_.Exception.Message) }
finally {
  try{ $doc.Close([ref]$true) }catch{ L('close fail: '+$_.Exception.Message) }
  try{ if($word.Documents.Count -eq 0){$word.Quit()} }catch{}
  try{ [Runtime.InteropServices.Marshal]::ReleaseComObject($word)|Out-Null }catch{}
}
Write-Output 'TABLE_DONE'
