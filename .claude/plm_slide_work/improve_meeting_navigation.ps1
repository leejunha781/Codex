$ErrorActionPreference="Stop"
$dir="D:\이력서\AVEVA - Marine Principal Technical Support & Consultant – PLM SME, Busan\Proposal"
$deck=Join-Path $dir "Future_Industrial_PLM_Meeting_Deck_EN_Rebuilt_v2.pptx"
$stamp=Get-Date -Format "yyyyMMdd_HHmmss"
$backupDir=Join-Path $dir ("_backup_"+$stamp)
$backup=Join-Path $backupDir ("Future_Industrial_PLM_Meeting_Deck_EN_Rebuilt_v2_BEFORE_MEETING_NAV_"+$stamp+".pptx")
$render="C:\Users\namma\.claude\plm_slide_work\meeting_current\slide02_after.png"
New-Item -ItemType Directory -Force -Path $backupDir|Out-Null
Copy-Item -LiteralPath $deck -Destination $backup -Force

function Rgb([string]$h){$h=$h.TrimStart('#');[int]([Convert]::ToInt32($h.Substring(0,2),16)+[Convert]::ToInt32($h.Substring(2,2),16)*256+[Convert]::ToInt32($h.Substring(4,2),16)*65536)}
$ppt=$null;$pres=$null
try{
 $ppt=New-Object -ComObject PowerPoint.Application
 $pres=$ppt.Presentations.Open($deck,$false,$false,$false)
 $sl=$pres.Slides.Item(2)
 for($i=$sl.Shapes.Count;$i-ge1;$i--){if($sl.Shapes.Item($i).Name-like"MeetingNavRange*"){$sl.Shapes.Item($i).Delete()}}
 foreach($sh in $sl.Shapes){
  try{
   if($sh.HasTextFrame-eq-1-and$sh.TextFrame.HasText-eq-1-and$sh.TextFrame.TextRange.Text-like"Abbreviations*Glossary*"){
    $sh.TextFrame.TextRange.Text="Abbreviations → Glossary appendix, p.31–33"
   }
  }catch{}
 }
 $ranges=@(
  @{t=102;r="p.3–5";c="31C5F4"},@{t=154;r="p.6";c="4F8BE8"},@{t=206;r="p.7–12";c="E84AA8"},
  @{t=258;r="p.13–25";c="49C28C"},@{t=310;r="p.26";c="F0832F"},@{t=362;r="p.27–29";c="E8B23A"},@{t=414;r="p.30";c="9B7BE0"}
 )
 $k=1
 foreach($d in $ranges){
  $p=$sl.Shapes.AddShape(5,[single]858,[single]($d.t+12),[single]50,[single]17)
  $p.Name=("MeetingNavRange"+$k);$p.Fill.Solid();$p.Fill.ForeColor.RGB=[int](Rgb "102033");$p.Line.Visible=-1;$p.Line.ForeColor.RGB=[int](Rgb $d.c);$p.Line.Weight=[single]0.7
  $p.TextFrame.MarginLeft=0;$p.TextFrame.MarginRight=0;$p.TextFrame.MarginTop=1;$p.TextFrame.MarginBottom=0;$p.TextFrame.VerticalAnchor=3
  $tr=$p.TextFrame.TextRange;$tr.Text=$d.r;$tr.Font.Name="Aptos";$tr.Font.Size=[single]7.5;$tr.Font.Bold=[int]1;$tr.Font.Color.RGB=[int](Rgb $d.c);$tr.ParagraphFormat.Alignment=2
  $k++
 }
 $pres.Save();$sl.Export($render,"PNG",1600,900)
}finally{
 if($pres){try{$pres.Close()}catch{}}
 if($ppt){try{if($ppt.Presentations.Count-eq0){$ppt.Quit()}}catch{}}
}
Write-Output("BACKUP="+$backup);Write-Output("RENDER="+$render)
