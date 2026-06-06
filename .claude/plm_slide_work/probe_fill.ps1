$ErrorActionPreference = "Stop"
$src = "D:\이력서\AVEVA - Marine Principal Technical Support & Consultant – PLM SME, Busan\Proposal\Future_Industrial_PLM_Meeting_Deck_EN_Rebuilt_v2.pptx"
$rdir = "C:\Users\namma\.claude\plm_slide_work\render_v2"
$ppt = New-Object -ComObject PowerPoint.Application
$pres = $ppt.Presentations.Open($src, $true, $false, $false)

function FillInfo($idx, $names){
  $s = $pres.Slides.Item($idx)
  Write-Output "---- SLIDE $idx ----"
  foreach($sh in $s.Shapes){
    if ($names -contains $sh.Name){
      $ft=""; $fc=""; $tr=""; $lv=""
      try { $ft=$sh.Fill.Type } catch {}
      try { $fc=$sh.Fill.ForeColor.RGB } catch {}
      try { $tr=[math]::Round($sh.Fill.Transparency,3) } catch {}
      try { $lv=$sh.Line.Visible } catch {}
      Write-Output ("  {0,-22} fillType={1} foreRGB={2} transp={3} lineVis={4}" -f $sh.Name,$ft,$fc,$tr,$lv)
    }
  }
}
FillInfo 5  @("Text 4","Text 5","Text 6","Text 7","Text 8","Text 9")
FillInfo 7  @("Text 4","Text 5","Text 6")
FillInfo 16 @("Picture 2","Rectangle 3","Rounded Rectangle 4","TextBox 5")
FillInfo 30 @("Rounded Rectangle 2","Rounded Rectangle 11","Rounded Rectangle 28","Rounded Rectangle 40","Rounded Rectangle 63")

# high-res re-render of the four
foreach($i in 5,7,16,30){
  $pres.Slides.Item($i).Export((Join-Path $rdir ("hi{0:D2}.png" -f $i)), "PNG", 1600, 900)
}
$pres.Close()
if ($ppt.Presentations.Count -eq 0){ $ppt.Quit() }
Write-Output "FILL PROBE + RENDER DONE"
