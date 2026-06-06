$ErrorActionPreference="Stop"
$deck="D:\이력서\AVEVA - Marine Principal Technical Support & Consultant – PLM SME, Busan\Proposal\Future_Industrial_PLM_Meeting_Deck_EN_Rebuilt_v2.pptx"
$render="C:\Users\namma\.claude\plm_slide_work\meeting_current\render_final"
$report="C:\Users\namma\.claude\plm_slide_work\meeting_current\verification.txt"
New-Item -ItemType Directory -Force -Path $render|Out-Null
Get-ChildItem -LiteralPath $render -Filter *.png -ErrorAction SilentlyContinue|Remove-Item -Force
$ppt=$null;$pres=$null;$tables=0;$pics=0;$nav=0;$gloss=""
try{
 $ppt=New-Object -ComObject PowerPoint.Application;$pres=$ppt.Presentations.Open($deck,$true,$false,$false)
 foreach($sl in $pres.Slides){
  foreach($sh in $sl.Shapes){
   if($sh.Type-eq13){$pics++};try{if($sh.HasTable){$tables++}}catch{}
   if($sh.Name-like"MeetingNavRange*"){$nav++}
   try{if($sl.SlideIndex-eq2-and$sh.HasTextFrame-eq-1-and$sh.TextFrame.HasText-eq-1-and$sh.TextFrame.TextRange.Text-like"Abbreviations*"){$gloss=$sh.TextFrame.TextRange.Text}}catch{}
  }
  $sl.Export((Join-Path $render ("slide-{0:D2}.png"-f$sl.SlideIndex)),"PNG",1280,720)
 }
 $txt="SLIDES=$($pres.Slides.Count)`r`nTABLES=$tables`r`nPICTURES=$pics`r`nMEETING_NAV_CHIPS=$nav`r`nGLOSSARY_POINTER=$gloss"
 [IO.File]::WriteAllText($report,$txt,[Text.Encoding]::UTF8)
}finally{if($pres){try{$pres.Close()}catch{}};if($ppt){try{if($ppt.Presentations.Count-eq0){$ppt.Quit()}}catch{}}}
try{$fs=[IO.File]::Open($deck,'Open','ReadWrite','None');$fs.Close();$lock="unlocked"}catch{$lock="LOCKED"}
Write-Output("LOCK="+$lock);Write-Output("REPORT="+$report);Write-Output("RENDER="+$render)
