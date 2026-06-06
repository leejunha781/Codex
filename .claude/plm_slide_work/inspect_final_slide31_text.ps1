$ErrorActionPreference = 'Stop'
$ppt = 'D:\이력서\AVEVA - Marine Principal Technical Support & Consultant – PLM SME, Busan\Proposal\Future_Industrial_PLM_Meeting_Deck_EN_Final.pptx'
$pp=$null; $pres=$null
try {
  $pp=New-Object -ComObject PowerPoint.Application
  $pp.Visible=-1
  $pres=$pp.Presentations.Open($ppt,$false,$true,$false)
  $s=$pres.Slides.Item(31)
  foreach($name in @('TextBox 14','TextBox 53','TextBox 56')){
    foreach($sh in @($s.Shapes)){
      if($sh.Name -eq $name){
        $txt=$sh.TextFrame.TextRange.Text -replace "`r|`n", ' '
        "$name = $txt"
      }
    }
  }
}
finally { if($pres){$pres.Close()}; if($pp){$pp.Quit()} }
