$ErrorActionPreference = "Stop"
$src = "D:\이력서\AVEVA - Marine Principal Technical Support & Consultant – PLM SME, Busan\Proposal\Future_Industrial_PLM_Meeting_Deck_EN_Rebuilt_v2.pptx"
$img08 = "C:\Users\namma\.claude\plm_slide_work\img08.png"
$img16 = "C:\Users\namma\.claude\plm_slide_work\img16.png"
$rdir  = "C:\Users\namma\.claude\plm_slide_work\render_v2"
$SCNAVY = 2496011  # #0B1626

$ppt = New-Object -ComObject PowerPoint.Application
$pres = $ppt.Presentations.Open($src, $false, $false, $false)
function Retry($b){ for($k=0;$k -lt 6;$k++){ try{ & $b; return }catch{ Start-Sleep -Milliseconds 120 } }; & $b }
function GetShape($sl,$n){ foreach($sh in $sl.Shapes){ if($sh.Name -eq $n){ return $sh } }; return $null }

# ---- Slide 8: img08 as background + scrim behind opaque panels ----
$s8 = $pres.Slides.Item(8)
$s8.FollowMasterBackground = $false
Retry { $s8.Background.Fill.UserPicture($img08) }
$scr = $s8.Shapes.AddShape(1, 0, 0, 960, 540)
$scr.Fill.Solid(); Retry { $scr.Fill.ForeColor.RGB=[int]$SCNAVY }; Retry { $scr.Fill.Transparency=[single]0.35 }
$scr.Line.Visible = 0
$scr.Name = "BgScrim8"
$scr.ZOrder(1)   # msoSendToBack -> above bg fill, below all content
Write-Output "slide8: bg image + scrim set"

# ---- Slide 16: swap generic Picture 2 -> img16 in the VISUAL PROOF panel ----
$s16 = $pres.Slides.Item(16)
$p2 = GetShape $s16 "Picture 2"
if($p2){
  $L=$p2.Left; $T=$p2.Top; $W=$p2.Width; $H=$p2.Height
  $p2.Delete()
  $np = $s16.Shapes.AddPicture($img16, 0, -1, $L, $T, $W, $H)
  $np.Name = "Picture 2"
  $np.ZOrder(1)   # send behind the panel scrim + title/caption
  Write-Output ("slide16: img16 placed at L=$([int]$L) T=$([int]$T) W=$([int]$W) H=$([int]$H)")
} else { Write-Output "slide16: Picture 2 NOT FOUND" }
# lighten the panel scrim so img16 shows (was 0.15 = very dark)
$r3 = GetShape $s16 "Rectangle 3"
if($r3){ Retry { $r3.Fill.Transparency=[single]0.40 }; Write-Output ("slide16 Rectangle 3 transp -> " + [math]::Round($r3.Fill.Transparency,3)) }

$pres.Save()
foreach($i in 8,16){ $pres.Slides.Item($i).Export((Join-Path $rdir ("img_edit{0:D2}.png" -f $i)), "PNG", 1600, 900) }
$cnt = $pres.Slides.Count
$pres.Close()
if($ppt.Presentations.Count -eq 0){ $ppt.Quit() }
Write-Output ("IMAGE EDIT DONE. slides=" + $cnt)
