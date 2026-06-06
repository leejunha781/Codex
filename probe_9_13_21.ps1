$ErrorActionPreference = "Stop"
$prop = "D:\이력서\AVEVA - Marine Principal Technical Support & Consultant – PLM SME, Busan\Proposal\plm_integrated_english_package\04_Presentation Image"
$dst = "C:\Users\namma\.claude\plm_slide_work"
$cp = @{
 "05_a_high_detail_isometric_3d_infographic_scene_a_la.png"="img05.png"
 "06_a_highly_detailed_clean_futuristic_infographic.png"="img06.png"
 "07_additional_architecture_visual.png"="img07.png"
}
foreach($k in $cp.Keys){ Copy-Item (Join-Path $prop $k) (Join-Path $dst $cp[$k]) -Force }
"copied 05/06/07"

$src = "D:\이력서\AVEVA - Marine Principal Technical Support & Consultant – PLM SME, Busan\Proposal\Future_Industrial_PLM_Meeting_Deck_EN_Rebuilt_v2.pptx"
$ppt = New-Object -ComObject PowerPoint.Application
$pres = $ppt.Presentations.Open($src, $true, $false, $false)
function Dump($idx){
  $s=$pres.Slides.Item($idx)
  "==== SLIDE $idx (shapes=$($s.Shapes.Count)) bgFill=$($s.Background.Fill.Type) ===="
  foreach($sh in $s.Shapes){
    $k="shp$($sh.Type)"; if($sh.Type -eq 13){$k="PICTURE"}
    $tp=""; if($sh.Type -ne 13){ try{ if($sh.Fill.Type -eq 1){$tp="transp="+[math]::Round($sh.Fill.Transparency,2)} }catch{} }
    $t=""; try{ if($sh.HasTextFrame -eq -1 -and $sh.TextFrame.HasText -eq -1){$t=$sh.TextFrame.TextRange.Text} }catch{}
    $t=($t -replace "[\r\n]"," "); if($t.Length -gt 28){$t=$t.Substring(0,28)}
    "z={0,2} {1,-8} {2,-20} L={3,4} T={4,4} W={5,4} H={6,4} {7,-11} '{8}'" -f $sh.ZOrderPosition,$k,$sh.Name,[int]$sh.Left,[int]$sh.Top,[int]$sh.Width,[int]$sh.Height,$tp,$t
  }
}
Dump 9
Dump 13
Dump 21
$pres.Close()
if($ppt.Presentations.Count -eq 0){$ppt.Quit()}
