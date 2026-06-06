$ErrorActionPreference = "Stop"
$src = "D:\이력서\AVEVA - Marine Principal Technical Support & Consultant – PLM SME, Busan\Proposal\Future_Industrial_PLM_Meeting_Deck_EN_Rebuilt_v2.pptx"
$ppt = New-Object -ComObject PowerPoint.Application
$pres = $ppt.Presentations.Open($src, $true, $false, $false)
$s = $pres.Slides.Item(18)
"SLIDE 18 shapes=$($s.Shapes.Count)  bgFill=$($s.Background.Fill.Type)"
# show shapes that are large (potential background/app-window panels) with fill info
foreach($sh in $s.Shapes){
  $area = $sh.Width * $sh.Height
  if($area -gt 90000 -or $sh.Type -eq 13){
    $ft="";$tp=""
    try{$ft=$sh.Fill.Type}catch{}
    try{ if($sh.Fill.Type -eq 1){$tp=[math]::Round($sh.Fill.Transparency,2)} }catch{}
    $k="shp$($sh.Type)"; if($sh.Type -eq 13){$k="PICTURE"}
    "z={0,3} {1,-9} {2,-22} L={3,4} T={4,4} W={5,4} H={6,4} fillT={7} transp={8}" -f $sh.ZOrderPosition,$k,$sh.Name,[int]$sh.Left,[int]$sh.Top,[int]$sh.Width,[int]$sh.Height,$ft,$tp
  }
}
$pres.Close()
if($ppt.Presentations.Count -eq 0){$ppt.Quit()}
