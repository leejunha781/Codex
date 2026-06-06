$ErrorActionPreference="Stop"
$deck="D:\이력서\AVEVA - Marine Principal Technical Support & Consultant – PLM SME, Busan\Proposal\Future_Industrial_PLM_Meeting_Deck_EN_Rebuilt_v2.pptx"
$ppt=$null;$pres=$null
try{
 $ppt=New-Object -ComObject PowerPoint.Application;$pres=$ppt.Presentations.Open($deck,$true,$false,$false)
 Write-Output("SLIDES="+$pres.Slides.Count)
 foreach($idx in @(3,26)){
  $tables=0;$tableCells=0
  foreach($sh in $pres.Slides.Item($idx).Shapes){
   try{if($sh.HasTable){$tables++;$tableCells+=($sh.Table.Rows.Count*$sh.Table.Columns.Count)}}catch{}
  }
  Write-Output(("SLIDE_{0}_TABLES={1}"-f$idx,$tables))
  Write-Output(("SLIDE_{0}_TABLE_CELLS={1}"-f$idx,$tableCells))
 }
}finally{if($pres){$pres.Close()};if($ppt){if($ppt.Presentations.Count-eq0){$ppt.Quit()}}}
try{$fs=[IO.File]::Open($deck,'Open','ReadWrite','None');$fs.Close();Write-Output "LOCK=unlocked"}catch{Write-Output "LOCK=LOCKED"}
