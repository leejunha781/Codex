$ErrorActionPreference = "Stop"
$ppt = "D:\이력서\AVEVA - Marine Principal Technical Support & Consultant – PLM SME, Busan\Proposal\Future_Industrial_PLM_Meeting_Deck_EN.pptx"
$pp = $null; $pres = $null
function Dump-Slide($slide){
  "SLIDE $($slide.SlideIndex): $($slide.Shapes.Count) shapes"
  foreach($sh in @($slide.Shapes)){
    $txt=''
    try { if($sh.HasTextFrame -and $sh.TextFrame.HasText){ $txt = $sh.TextFrame.TextRange.Text } } catch {}
    if($txt.Length -gt 80){ $txt = $txt.Substring(0,80) }
    $txt = $txt -replace "`r|`n"," "
    "{0}`t{1}`tL={2:N1} T={3:N1} W={4:N1} H={5:N1}`t{6}" -f $sh.Id,$sh.Name,$sh.Left,$sh.Top,$sh.Width,$sh.Height,$txt
  }
}
try {
  $pp = New-Object -ComObject PowerPoint.Application
  $pp.Visible = -1
  $pres = $pp.Presentations.Open($ppt,$false,$false,$false)
  Dump-Slide $pres.Slides.Item(13)
  '---'
  Dump-Slide $pres.Slides.Item(14)
}
finally {
  if($pres){ $pres.Close() }
  if($pp){ $pp.Quit() }
}
