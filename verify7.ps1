$ErrorActionPreference = "Stop"
$v2 = "D:\이력서\AVEVA - Marine Principal Technical Support & Consultant – PLM SME, Busan\Proposal\Future_Industrial_PLM_Meeting_Deck_EN_Rebuilt_v2.pptx"
$name = "Future_Industrial_PLM_Meeting_Deck_EN_Rebuilt_v2.pptx"
$ppt = New-Object -ComObject PowerPoint.Application
$pres=$null
foreach($p in @($ppt.Presentations)){ if($p.Name -eq $name){ $pres=$p; break } }
$byUs=$false
if($pres -eq $null){ $pres=$ppt.Presentations.Open($v2,-1,0,0); $byUs=$true }
Write-Output ("TOTAL SLIDES = " + $pres.Slides.Count)
for($i=10;$i -le 17;$i++){
  $s=$pres.Slides.Item($i); $title=""
  foreach($sh in $s.Shapes){ try{ if($sh.HasTextFrame -eq -1 -and $sh.TextFrame.HasText -eq -1 -and $sh.Top -lt 40 -and $sh.Left -lt 60){ $title=$sh.TextFrame.TextRange.Text.Trim(); break } }catch{} }
  if($title -eq ""){ foreach($sh in $s.Shapes){ try{ if($sh.HasTextFrame -eq -1 -and $sh.TextFrame.HasText -eq -1){ $title=$sh.TextFrame.TextRange.Text.Trim(); break } }catch{} } }
  if($title.Length -gt 64){ $title=$title.Substring(0,64) }
  Write-Output ("  " + $i + ": " + $title)
}
if($byUs){ $pres.Close() }
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($pres) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($ppt) | Out-Null
[GC]::Collect()
Write-Output "VERIFY_DONE"
