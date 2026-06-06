$ErrorActionPreference = "Stop"
$ppt = "D:\이력서\AVEVA - Marine Principal Technical Support & Consultant – PLM SME, Busan\Proposal\Future_Industrial_PLM_Meeting_Deck_EN_V6.pptx"
$pp = $null; $pres = $null
function Dump-Named($slide,$names){ foreach($name in $names){ foreach($sh in @($slide.Shapes)){ if($sh.Name -eq $name){ $txt=''; try { if($sh.HasTextFrame -and $sh.TextFrame.HasText){ $txt=$sh.TextFrame.TextRange.Text } } catch {}; $txt=$txt -replace "`r|`n"," "; if($txt.Length -gt 60){$txt=$txt.Substring(0,60)}; "{0}`tL={1:N1} T={2:N1} W={3:N1} H={4:N1}`t{5}" -f $name,$sh.Left,$sh.Top,$sh.Width,$sh.Height,$txt } } }
}
try {
  $pp = New-Object -ComObject PowerPoint.Application
  $pp.Visible = -1
  $pres = $pp.Presentations.Open($ppt,$false,$true,$false)
  'SLIDE13'; Dump-Named $pres.Slides.Item(13) @('Picture 96','Rounded Rectangle 53','TextBox 59','TextBox 60','Rounded Rectangle 61','TextBox 78','TextBox 79','Rounded Rectangle 80','TextBox 82')
  'SLIDE14'; Dump-Named $pres.Slides.Item(14) @('Picture 1','Rounded Rectangle 57','TextBox 74','Rounded Rectangle 75','TextBox 94')
}
finally { if($pres){$pres.Close()}; if($pp){$pp.Quit()} }
