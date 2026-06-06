$ErrorActionPreference = "Stop"
$dir = "D:\이력서\AVEVA - Marine Principal Technical Support & Consultant – PLM SME, Busan\Proposal"
$v2 = Join-Path $dir "Future_Industrial_PLM_Meeting_Deck_EN_Rebuilt_v2.pptx"
$ko = Join-Path $dir "Future_Industrial_PLM_Meeting_Deck_KO_Rebuilt_v2.pptx"
$koName = "Future_Industrial_PLM_Meeting_Deck_KO_Rebuilt_v2.pptx"

Copy-Item -LiteralPath $v2 -Destination $ko -Force
Write-Output ("KO_FILE=" + $ko)

$ppt = New-Object -ComObject PowerPoint.Application
$pres=$null
foreach($p in @($ppt.Presentations)){ if($p.Name -eq $koName){ $pres=$p; break } }
$byUs=$false
if($pres -eq $null){ $pres=$ppt.Presentations.Open($ko,-1,0,0); $byUs=$true }
Write-Output ("SLIDES=" + $pres.Slides.Count)

function DumpShape($sh, $sidx){
  try{
    if($sh.Type -eq 6){ foreach($c in $sh.GroupItems){ DumpShape $c $sidx }; return }
    if($sh.HasTable -eq -1){ $tb=$sh.Table; for($r=1;$r -le $tb.Rows.Count;$r++){ for($c=1;$c -le $tb.Columns.Count;$c++){ $t=$tb.Cell($r,$c).Shape.TextFrame.TextRange.Text; if($t.Trim().Length -gt 0){ $k=($t -replace "[\r\n]","¶"); Write-Output ("S$sidx`t$k") } } }; return }
    if($sh.HasTextFrame -eq -1 -and $sh.TextFrame.HasText -eq -1){ $t=$sh.TextFrame.TextRange.Text; if($t.Trim().Length -gt 0){ $k=($t -replace "[\r\n]","¶"); Write-Output ("S$sidx`t$k") } }
  }catch{}
}
for($i=1;$i -le $pres.Slides.Count;$i++){ foreach($sh in $pres.Slides.Item($i).Shapes){ DumpShape $sh $i } }

if($byUs){ $pres.Close() }
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($pres) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($ppt) | Out-Null
[GC]::Collect()
Write-Output "EXTRACT_DONE"
