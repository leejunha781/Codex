$ErrorActionPreference="Stop"
$deck="D:\이력서\AVEVA - Marine Principal Technical Support & Consultant – PLM SME, Busan\Proposal\Future_Industrial_PLM_Meeting_Deck_EN_Rebuilt_v2.pptx"
$out="C:\Users\namma\.claude\plm_slide_work\meeting_current\comparison_probe.txt"
$ppt=$null;$pres=$null;$sb=New-Object Text.StringBuilder
try{
 $ppt=New-Object -ComObject PowerPoint.Application;$pres=$ppt.Presentations.Open($deck,$true,$false,$false)
 foreach($idx in @(3,26)){
  $sl=$pres.Slides.Item($idx);[void]$sb.AppendLine(("===== SLIDE "+$idx+" ====="))
  foreach($sh in $sl.Shapes){
   $t="";try{if($sh.HasTextFrame-eq-1-and$sh.TextFrame.HasText-eq-1){$t=($sh.TextFrame.TextRange.Text-replace"[\r\n\v]"," | ").Trim()}}catch{}
   [void]$sb.AppendLine(("z={0,3} type={1,2} {2,-24} L={3,5:N0} T={4,5:N0} W={5,5:N0} H={6,5:N0} :: {7}"-f$sh.ZOrderPosition,$sh.Type,$sh.Name,$sh.Left,$sh.Top,$sh.Width,$sh.Height,$t))
  }
 }
}finally{if($pres){$pres.Close()};if($ppt){if($ppt.Presentations.Count-eq0){$ppt.Quit()}}}
[IO.File]::WriteAllText($out,$sb.ToString(),[Text.Encoding]::UTF8);Write-Output("OUT="+$out)
