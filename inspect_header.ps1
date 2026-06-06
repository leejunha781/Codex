$ErrorActionPreference = 'Continue'
$out = New-Object System.Text.StringBuilder
function L($m){ [void]$out.AppendLine($m) }

# PDF lock test
$pdf = "D:\이력서\ITT Cannon\Joonha_Lee_ITT_Cannon_FAE_Resume_Final_Integrated_v2.pdf"
try { $fs=[System.IO.File]::Open($pdf,'Open','ReadWrite','None'); $fs.Close(); L('PDF: unlocked (can refresh)') } catch { L('PDF: LOCKED') }
$acr = (Get-Process Acrobat,AcroRd32 -ErrorAction SilentlyContinue | Measure-Object).Count
L("Acrobat procs: $acr")

$word = New-Object -ComObject Word.Application
$word.Visible = $false
try { $word.DisplayAlerts = 0 } catch {}
$path = "D:\이력서\ITT Cannon\Joonha_Lee_ITT_Cannon_FAE_Resume_Final_Integrated_v2.docx"
$doc = $word.Documents.Open($path, $false, $true)

L("InlineShapes total: " + $doc.InlineShapes.Count)
$si=0
foreach($s in $doc.InlineShapes){
  $si++
  $inT=$false; try{$inT=$s.Range.Information(12)}catch{}
  $rn=0;$cn=0
  try{ $rn=$s.Range.Information(13) }catch{}   # wdStartOfRangeRowNumber=13
  try{ $cn=$s.Range.Information(16) }catch{}   # wdStartOfRangeColumnNumber=16
  L(("InlineShape #{0}: type={1} w={2} h={3} inTable={4} row={5} col={6}" -f $si,$s.Type,[math]::Round($s.Width,0),[math]::Round($s.Height,0),$inT,$rn,$cn))
}
L("Floating Shapes total: " + $doc.Shapes.Count)

# Table #1 structure
$t1=$doc.Tables.Item(1)
L("Table#1 rows="+$t1.Rows.Count+" cols="+$t1.Columns.Count)
foreach($c in $t1.Range.Cells){
  $txt=(($c.Range.Text) -replace "[`r`n`a]"," ").Trim(); if($txt.Length -gt 40){$txt=$txt.Substring(0,40)}
  $hasImg=$c.Range.InlineShapes.Count
  L(("  cell r="+$c.RowIndex+" c="+$c.ColumnIndex+" img="+$hasImg+" | "+$txt))
}

# confirm Values section + Requirement table presence
$hasV=$false;$hasReqH=$false
foreach($p in $doc.Paragraphs){ $t=($p.Range.Text -replace "[`r`n`a]","").Trim(); if($t -eq 'Alignment with ITT Purpose, Values & DNA'){$hasV=$true}; if($t -eq 'Attachment: Condensed ITT Cannon FAE Requirement Match'){$hasReqH=$true} }
L("Values heading present: $hasV ; Requirement heading present: $hasReqH ; total tables: "+$doc.Tables.Count)

Set-Content -Path "C:\Users\namma\.claude\itt_work\header_info.txt" -Value $out.ToString() -Encoding UTF8
$doc.Close([ref]$false)
if($word.Documents.Count -eq 0){$word.Quit()}
[Runtime.InteropServices.Marshal]::ReleaseComObject($word)|Out-Null
Write-Output "HEADER_INSPECT_DONE"
