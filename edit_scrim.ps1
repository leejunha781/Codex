$ErrorActionPreference = "Stop"
$src = "D:\이력서\AVEVA - Marine Principal Technical Support & Consultant – PLM SME, Busan\Proposal\Future_Industrial_PLM_Meeting_Deck_EN_Rebuilt_v2.pptx"
$CYAN = 15646772
$ARR  = [char]0x2192   # arrow
$EN   = [char]0x2013   # en dash

$ppt = New-Object -ComObject PowerPoint.Application
$pres = $ppt.Presentations.Open($src, $false, $false, $false)

function Retry($block){ for($k=0;$k -lt 6;$k++){ try{ & $block; return }catch{ Start-Sleep -Milliseconds 120 } }; & $block }
function GetShape($slide,$name){ foreach($sh in $slide.Shapes){ if($sh.Name -eq $name){ return $sh } }; return $null }

# ---- (1) Slide 2: glossary pointer callout (additive; existing shapes untouched) ----
$s2 = $pres.Slides.Item(2)
$ptr = $s2.Shapes.AddTextbox(1, 540, 55, 334, 18)
$tf = $ptr.TextFrame; $tf.WordWrap=-1; $tf.AutoSize=0
$tf.MarginLeft=0;$tf.MarginRight=0;$tf.MarginTop=0;$tf.MarginBottom=0
$tr = $tf.TextRange
$tr.Text = "Abbreviations " + $ARR + " Glossary appendix, p.31" + $EN + "32"
$tr.Font.Name="Aptos"; $tr.Font.Size=[single]9; $tr.Font.Bold=[int]0
Retry { $tr.Font.Color.RGB=[int]$CYAN }
$tr.ParagraphFormat.Alignment=[int]3   # right
$ptr.Name = "GlossaryPointer"
Write-Output "slide2 pointer added"

# ---- (2) Slide 16: deepen existing scrim over the photo ----
$s16 = $pres.Slides.Item(16)
$r3 = GetShape $s16 "Rectangle 3"
if ($r3){ Retry { $r3.Fill.Transparency = [single]0.15 }; Write-Output ("slide16 Rectangle 3 transp -> " + [math]::Round($r3.Fill.Transparency,3)) }
else { Write-Output "slide16 Rectangle 3 NOT FOUND" }

# ---- (3) Slide 30: deepen full-bleed background scrim ----
$s30 = $pres.Slides.Item(30)
$r2 = GetShape $s30 "Rounded Rectangle 2"
if ($r2){ Retry { $r2.Fill.Transparency = [single]0.30 }; Write-Output ("slide30 Rounded Rectangle 2 transp -> " + [math]::Round($r2.Fill.Transparency,3)) }
else { Write-Output "slide30 Rounded Rectangle 2 NOT FOUND" }

$pres.Save()
$cnt = $pres.Slides.Count
# re-render the three touched slides hi-res
$rdir = "C:\Users\namma\.claude\plm_slide_work\render_v2"
foreach($i in 2,16,30){ $pres.Slides.Item($i).Export((Join-Path $rdir ("edit{0:D2}.png" -f $i)), "PNG", 1600, 900) }
$pres.Close()
if ($ppt.Presentations.Count -eq 0){ $ppt.Quit() }
Write-Output ("EDIT DONE. slides=" + $cnt)
