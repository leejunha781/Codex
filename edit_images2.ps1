$ErrorActionPreference = "Stop"
$src = "D:\이력서\AVEVA - Marine Principal Technical Support & Consultant – PLM SME, Busan\Proposal\Future_Industrial_PLM_Meeting_Deck_EN_Rebuilt_v2.pptx"
$img05 = "C:\Users\namma\.claude\plm_slide_work\img05.png"
$img06 = "C:\Users\namma\.claude\plm_slide_work\img06.png"
$img16 = "C:\Users\namma\.claude\plm_slide_work\img16.png"
$rdir  = "C:\Users\namma\.claude\plm_slide_work\render_v2"
$SCNAVY = 2496011

$ppt = New-Object -ComObject PowerPoint.Application
$pres = $ppt.Presentations.Open($src, $false, $false, $false)
function Retry($b){ for($k=0;$k -lt 6;$k++){ try{ & $b; return }catch{ Start-Sleep -Milliseconds 120 } }; & $b }
function GetShape($sl,$n){ foreach($sh in $sl.Shapes){ if($sh.Name -eq $n){ return $sh } }; return $null }
function SwapPicture($sl,$name,$path){
  $p = GetShape $sl $name
  if(-not $p){ Write-Output "  $name NOT FOUND"; return }
  $L=$p.Left;$T=$p.Top;$W=$p.Width;$H=$p.Height
  $p.Delete()
  $np = $sl.Shapes.AddPicture($path,0,-1,$L,$T,$W,$H)
  $np.Name = $name
  $np.ZOrder(1)
  Write-Output ("  swapped $name -> img at L=$([int]$L) T=$([int]$T) W=$([int]$W) H=$([int]$H)")
}

# ---- Slide 9: Picture 5 -> img06 (convergence/brain), keep existing scrim ----
"slide9:"
SwapPicture $pres.Slides.Item(9) "Picture 5" $img06

# ---- Slide 13: Picture 13 -> img05 (layered stack) ----
"slide13:"
SwapPicture $pres.Slides.Item(13) "Picture 13" $img05

# ---- Slide 21: img16 loop as background + scrim ----
"slide21:"
$s21 = $pres.Slides.Item(21)
$s21.FollowMasterBackground = $false
Retry { $s21.Background.Fill.UserPicture($img16) }
$scr = $s21.Shapes.AddShape(1,0,0,960,540)
$scr.Fill.Solid(); Retry { $scr.Fill.ForeColor.RGB=[int]$SCNAVY }; Retry { $scr.Fill.Transparency=[single]0.30 }
$scr.Line.Visible=0; $scr.Name="BgScrim21"; $scr.ZOrder(1)
"  bg img16 + scrim set"

$pres.Save()
foreach($i in 9,13,21){ $pres.Slides.Item($i).Export((Join-Path $rdir ("img2_{0:D2}.png" -f $i)), "PNG", 1600, 900) }
$cnt=$pres.Slides.Count
$pres.Close()
if($ppt.Presentations.Count -eq 0){$ppt.Quit()}
Write-Output ("DONE. slides=" + $cnt)
