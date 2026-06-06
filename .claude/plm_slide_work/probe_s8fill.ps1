$ErrorActionPreference = "Stop"
$src = "D:\이력서\AVEVA - Marine Principal Technical Support & Consultant – PLM SME, Busan\Proposal\Future_Industrial_PLM_Meeting_Deck_EN_Rebuilt_v2.pptx"
$ppt = New-Object -ComObject PowerPoint.Application
$pres = $ppt.Presentations.Open($src, $true, $false, $false)
$names = @("Rounded Rectangle 19","Rounded Rectangle 25","Rounded Rectangle 30","Rounded Rectangle 35","Rounded Rectangle 60","Rounded Rectangle 74","Rounded Rectangle 76","Oval 39","Rounded Rectangle 42","Rounded Rectangle 44")
$s = $pres.Slides.Item(8)
foreach($sh in $s.Shapes){
  if($names -contains $sh.Name){
    $ft="";$fc="";$tr=""
    try{$ft=$sh.Fill.Type}catch{}
    try{$fc=$sh.Fill.ForeColor.RGB}catch{}
    try{$tr=[math]::Round($sh.Fill.Transparency,3)}catch{}
    "{0,-22} fillType={1} foreRGB={2} transp={3}" -f $sh.Name,$ft,$fc,$tr
  }
}
$pres.Close()
if($ppt.Presentations.Count -eq 0){$ppt.Quit()}
