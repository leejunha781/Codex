$ErrorActionPreference='Stop'
$ppt='D:\이력서\AVEVA - Marine Principal Technical Support & Consultant – PLM SME, Busan\Proposal\Future_Industrial_PLM_Meeting_Deck_EN_Final.pptx'
$pp=$null; $pres=$null
try{
  $pp=New-Object -ComObject PowerPoint.Application
  $pp.Visible=-1
  $pres=$pp.Presentations.Open($ppt,$false,$true,$false)
  "SLIDES=$($pres.Slides.Count)"
  foreach($i in 1..$pres.Slides.Count){
    $title=''
    foreach($sh in @($pres.Slides.Item($i).Shapes)){
      try{if($sh.HasTextFrame -and $sh.TextFrame.HasText){$title=(($sh.TextFrame.TextRange.Text -replace "[`r`n]+",' ') -replace '\s+',' ').Trim(); break}}catch{}
    }
    if($title.Length -gt 80){$title=$title.Substring(0,80)}
    '[{0:00}] {1}' -f $i,$title
  }
}
finally{if($pres){$pres.Close()}; if($pp){$pp.Quit()}}
