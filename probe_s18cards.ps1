$ErrorActionPreference = "Stop"
$prop = "D:\이력서\AVEVA - Marine Principal Technical Support & Consultant – PLM SME, Busan\Proposal\plm_integrated_english_package\04_Presentation Image"
$dst = "C:\Users\namma\.claude\plm_slide_work"
Copy-Item (Join-Path $prop "04_a_wide_cinematic_industrial_technology_visualizat.png") (Join-Path $dst "img04.png") -Force
Copy-Item (Join-Path $prop "03_a_wide_cinematic_futuristic_industrial_engineerin.png") (Join-Path $dst "img03.png") -Force
"copied img03, img04"

$src = "D:\이력서\AVEVA - Marine Principal Technical Support & Consultant – PLM SME, Busan\Proposal\Future_Industrial_PLM_Meeting_Deck_EN_Rebuilt_v2.pptx"
$ppt = New-Object -ComObject PowerPoint.Application
$pres = $ppt.Presentations.Open($src, $true, $false, $false)
$s = $pres.Slides.Item(18)
# Sample fills of the 9 step cards (rounded rects ~132 wide) and the window
$cards = @("Rounded Rectangle 102","Rounded Rectangle 110","Rounded Rectangle 118","Rounded Rectangle 126","Rounded Rectangle 134","Rounded Rectangle 142","Rounded Rectangle 150","Rounded Rectangle 158","Rounded Rectangle 166","Rectangle 6")
foreach($sh in $s.Shapes){
  if($cards -contains $sh.Name){
    $tp=""; try{ if($sh.Fill.Type -eq 1){$tp=[math]::Round($sh.Fill.Transparency,2)} }catch{}
    "{0,-22} L={1,4} T={2,4} W={3,4} H={4,4} fillType={5} transp={6}" -f $sh.Name,[int]$sh.Left,[int]$sh.Top,[int]$sh.Width,[int]$sh.Height,$sh.Fill.Type,$tp
  }
}
$pres.Close()
if($ppt.Presentations.Count -eq 0){$ppt.Quit()}
