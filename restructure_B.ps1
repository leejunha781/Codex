$ErrorActionPreference = 'Continue'
$logPath = "C:\Users\namma\.claude\itt_work\B_log.txt"
$log = New-Object System.Text.StringBuilder
function L($m){ [void]$log.AppendLine($m); Set-Content -Path $logPath -Value $log.ToString() -Encoding UTF8 }
function Norm($p){ return ($p.Range.Text -replace "[`r`n`a]","").Trim() }

$path = "D:\이력서\ITT Cannon\Joonha_Lee_ITT_Cannon_FAE_Resume_Final_Integrated_v2.docx"
$pdf  = "D:\이력서\ITT Cannon\Joonha_Lee_ITT_Cannon_FAE_Resume_Final_Integrated_v2.pdf"
$charcoal=4210752; $lightgray=15658734; $white=16777215

$word = New-Object -ComObject Word.Application
$word.Visible = $false
try { $word.DisplayAlerts = 0 } catch {}
$doc = $word.Documents.Open($path, $false, $false)
L('opened')

try {
  $t1 = $doc.Tables.Item(1)
  $oldImg = $t1.Cell(1,2).Range.InlineShapes.Count
  L("old header photo count=$oldImg")

  # insert new 6x3 table before the Professional Profile heading
  $prof=$null
  foreach($p in $doc.Paragraphs){ if((Norm $p) -eq 'Professional Profile and Interest in ITT Cannon FAE Role'){ $prof=$p; break } }
  if($prof -eq $null){ throw 'Profile heading not found' }
  $ar=$prof.Range.Duplicate; $ar.Collapse(1)
  $newT=$doc.Tables.Add($ar,6,3)
  L('new table added')

  # fill cells (before merges)
  $newT.Cell(1,1).Range.Text='Resume'
  $newT.Cell(2,1).Range.Text='Name';                 $newT.Cell(2,2).Range.Text='Joonha Lee'
  $newT.Cell(3,1).Range.Text='Position Applied For';  $newT.Cell(3,2).Range.Text='Field Application Engineer (FAE) - ITT Cannon Connector Korea'
  $newT.Cell(4,1).Range.Text='Location';              $newT.Cell(4,2).Range.Text='Gunpo-si, Gyeonggi-do, Korea'
  $newT.Cell(5,1).Range.Text='Phone';                 $newT.Cell(5,2).Range.Text='+82-10-2731-4581'
  $newT.Cell(6,1).Range.Text='Email';                 $newT.Cell(6,2).Range.Text='leejunha781@gmail.com'
  L('filled')

  # merge col3 rows 2-6 -> photo cell, then copy photo
  $newT.Cell(2,3).Merge($newT.Cell(6,3))
  $newT.Cell(2,3).Range.FormattedText = $t1.Cell(1,2).Range.FormattedText
  $imgN = $newT.Cell(2,3).Range.InlineShapes.Count
  L("photo copied; new photo cell images=$imgN")

  # merge row1 -> title
  $newT.Cell(1,1).Merge($newT.Cell(1,3))
  L('merged title row')

  # delete old table only if photo safely copied
  if($imgN -ge 1){ $t1.Delete(); L('old header deleted') } else { L('PHOTO COPY FAILED - old header kept') }

  # re-grab new table (now table #1)
  $h = $doc.Tables.Item(1)

  # styling
  $h.Range.Font.Name='Arial'
  $h.Range.ParagraphFormat.LineSpacingRule=0
  try{ $h.Borders.InsideLineStyle=[int]1; $h.Borders.OutsideLineStyle=[int]1; $h.Borders.InsideColor=[int]$lightgray; $h.Borders.OutsideColor=[int]$lightgray }catch{ L('border fail') }
  try{ $h.PreferredWidthType=2; $h.PreferredWidth=[single]100 }catch{}
  try{ $h.Columns.Item(1).PreferredWidthType=2; $h.Columns.Item(1).PreferredWidth=[single]24 }catch{ L('col1 width fail') }
  try{ $h.Columns.Item(2).PreferredWidthType=2; $h.Columns.Item(2).PreferredWidth=[single]56 }catch{ L('col2 width fail') }
  try{ $h.Columns.Item(3).PreferredWidthType=2; $h.Columns.Item(3).PreferredWidth=[single]20 }catch{ L('col3 width fail') }

  # title cell
  $tc=$h.Cell(1,1)
  $tc.Range.Font.Size=[single]18; $tc.Range.Font.Bold=[int]1; $tc.Range.Font.Color=[int]$white
  $tc.Range.ParagraphFormat.Alignment=[int]1   # center
  try{ $tc.Shading.BackgroundPatternColor=[int]$charcoal }catch{}
  try{ $tc.VerticalAlignment=[int]1 }catch{}

  # label + value cells
  for($r=2;$r -le 6;$r++){
    try{ $lc=$h.Cell($r,1); $lc.Range.Font.Size=[single]10; $lc.Range.Font.Bold=[int]1; $lc.Range.Font.Color=[int]$charcoal; $lc.Shading.BackgroundPatternColor=[int]$lightgray; $lc.VerticalAlignment=[int]1 }catch{ L("label r=$r fail") }
    try{ $vc=$h.Cell($r,2); $vc.Range.Font.Color=[int]$charcoal; $vc.VerticalAlignment=[int]1; if($r -eq 2){ $vc.Range.Font.Size=[single]13; $vc.Range.Font.Bold=[int]1 } else { $vc.Range.Font.Size=[single]11; $vc.Range.Font.Bold=[int]0 } }catch{ L("value r=$r fail") }
  }
  # photo cell center
  try{ $pc=$h.Cell(2,3); $pc.VerticalAlignment=[int]1; $pc.Range.ParagraphFormat.Alignment=[int]1 }catch{ L('photo cell align fail') }
  L('styled')

  # safety: bump any sub-10 (labels are 10; fine)
  $bump=0; foreach($p in $doc.Paragraphs){ $sz=$null;try{$sz=$p.Range.Font.Size}catch{}; if($sz -ne $null -and $sz -ne 9999999 -and $sz -lt 10){ $p.Range.Font.Size=[single]10; $bump++ } }
  L("bumped=$bump")

  $doc.Save(); L('saved')
  try{ $doc.ExportAsFixedFormat($pdf,17); L('pdf refreshed') }catch{ L('pdf fail: '+$_.Exception.Message) }
  L('PAGES='+$doc.ComputeStatistics(2)+' TABLES='+$doc.Tables.Count+' INLINESHAPES='+$doc.InlineShapes.Count)
}
catch { L('FATAL: '+$_.Exception.Message) }
finally {
  try{ $doc.Close([ref]$true) }catch{}
  try{ if($word.Documents.Count -eq 0){$word.Quit()} }catch{}
  try{ [Runtime.InteropServices.Marshal]::ReleaseComObject($word)|Out-Null }catch{}
}
Write-Output 'B_DONE'
