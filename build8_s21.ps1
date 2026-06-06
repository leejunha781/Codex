$ErrorActionPreference = "Stop"
$v2 = "D:\이력서\AVEVA - Marine Principal Technical Support & Consultant – PLM SME, Busan\Proposal\Future_Industrial_PLM_Meeting_Deck_EN_Rebuilt_v2.pptx"
$name = "Future_Industrial_PLM_Meeting_Deck_EN_Rebuilt_v2.pptx"
$outDir = "C:\Users\namma\.claude\plm_slide_work\review_png8"
function Rgb([string]$hex){ $hex=$hex.TrimStart('#'); [int]([Convert]::ToInt32($hex.Substring(0,2),16)+[Convert]::ToInt32($hex.Substring(2,2),16)*256+[Convert]::ToInt32($hex.Substring(4,2),16)*65536) }
function Retry($sb){ for($a=0;$a -lt 6;$a++){ try{ return (& $sb) }catch{ Start-Sleep -Milliseconds 80 } }; return (& $sb) }
$WHITE="FFFFFF"; $LIGHT="C7D6E5"; $MUTE="93A8C0"; $GREEN="49C28C"; $AMBER="E8B23A"; $RED="E5634E"
$BGD="0B1626"; $PANEL="14253C"; $GRID="2C4663"
$ppt = New-Object -ComObject PowerPoint.Application
foreach($p in @($ppt.Presentations)){ if($p.Name -eq $name){ try{ $p.Saved=$true; $p.Close() }catch{} } }
$pres=$ppt.Presentations.Open($v2,0,0,0)
function NoShadow($s){ try{$s.Shadow.Visible=0}catch{} }
function Panel($l,$t,$w,$h,$fill,$line,$lw,$trans,$rad){ $s=$script:slide.Shapes.AddShape(5,[single]$l,[single]$t,[single]$w,[single]$h); try{$s.Adjustments.Item(1)=[single]$rad}catch{}; $s.Fill.Solid(); Retry{$s.Fill.ForeColor.RGB=[int](Rgb $fill)}; $s.Fill.Transparency=[single]$trans; if($line){$s.Line.Visible=-1; Retry{$s.Line.ForeColor.RGB=[int](Rgb $line)}; $s.Line.Weight=[single]$lw}else{$s.Line.Visible=0}; NoShadow $s; return $s }
function Oval($l,$t,$w,$h,$fill){ $s=$script:slide.Shapes.AddShape(9,[single]$l,[single]$t,[single]$w,[single]$h); $s.Fill.Solid(); Retry{$s.Fill.ForeColor.RGB=[int](Rgb $fill)}; $s.Line.Visible=0; NoShadow $s; return $s }
function Txt($l,$t,$w,$h,$text,$size,$bold,$color,$font,$align,$anchor){ $tb=$script:slide.Shapes.AddTextbox(1,[single]$l,[single]$t,[single]$w,[single]$h); $tb.Fill.Visible=0;$tb.Line.Visible=0;NoShadow $tb; $tf=$tb.TextFrame;$tf.WordWrap=-1;$tf.AutoSize=0;$tf.MarginLeft=[single]2;$tf.MarginRight=[single]2;$tf.MarginTop=[single]1;$tf.MarginBottom=[single]1;$tf.VerticalAnchor=[int]$anchor; $tr=$tf.TextRange;$tr.Text=$text;$tr.Font.Size=[single]$size;$tr.Font.Bold=[int]$bold;$tr.Font.Name=[string]$font;Retry{$tr.Font.Color.RGB=[int](Rgb $color)};$tr.ParagraphFormat.Alignment=[int]$align; try{$tr.ParagraphFormat.SpaceBefore=[single]0;$tr.ParagraphFormat.SpaceAfter=[single]0}catch{}; return $tb }
function DelAll($sl){ foreach($sh in @($sl.Shapes)){ try{$sh.Delete()}catch{} } }
function FindLike($pres,$prefix){ foreach($sl in $pres.Slides){ foreach($sh in $sl.Shapes){ try{ if($sh.HasTextFrame -eq -1 -and $sh.TextFrame.HasText -eq -1){ if($sh.TextFrame.TextRange.Text.Trim().StartsWith($prefix)){ return $sl } } }catch{} } }; return $null }

$s = FindLike $pres "Main risks"
Write-Output ("s21 idx=" + $s.SlideIndex)
DelAll $s
$script:slide = $s
$bg=$script:slide.Shapes.AddShape(1,[single]0,[single]0,[single]960,[single]540); $bg.Line.Visible=0; NoShadow $bg; $bg.Fill.Solid(); Retry{$bg.Fill.ForeColor.RGB=[int](Rgb $BGD)}
Txt 26 14 760 30 "Main risks and mitigations" 26 -1 $WHITE "Aptos Display" 1 1 | Out-Null
Txt 770 16 158 14 ([string]$s.SlideIndex) 9 0 $MUTE "Aptos" 3 1 | Out-Null
Txt 26 50 760 16 "Likelihood x Impact — early agreement reduces rework" 11 0 $LIGHT "Aptos" 1 1 | Out-Null

# 3x3 matrix
$gx=128; $gy=120; $cw=96; $ch=86
$zoneF = @(@("3A2E10","3A1614","3A1614"),@("12352A","3A2E10","3A1614"),@("12352A","12352A","3A2E10"))
$zoneL = @(@($AMBER,$RED,$RED),@($GREEN,$AMBER,$RED),@($GREEN,$GREEN,$AMBER))
for($r=0;$r -lt 3;$r++){ for($c=0;$c -lt 3;$c++){ Panel ($gx+$c*$cw) ($gy+$r*$ch) ($cw-4) ($ch-4) $zoneF[$r][$c] $zoneL[$r][$c] 0.75 0 0.04 | Out-Null } }
# axis labels
$rl=@("High","Med","Low")
for($r=0;$r -lt 3;$r++){ Txt 78 ($gy+$r*$ch+30) 46 16 $rl[$r] 8.5 -1 $MUTE "Aptos" 3 1 | Out-Null }
$cl=@("Low","Med","High")
for($c=0;$c -lt 3;$c++){ Txt ($gx+$c*$cw) ($gy+3*$ch+4) ($cw-4) 14 $cl[$c] 8.5 -1 $MUTE "Aptos" 2 1 | Out-Null }
Txt 128 ($gy+3*$ch+22) 288 14 "Impact ->" 9 -1 $LIGHT "Aptos" 2 1 | Out-Null
Txt 40 100 90 14 "Likelihood" 9 -1 $LIGHT "Aptos" 1 1 | Out-Null
# risk dots (r,c) -> center
function DotAt($r,$c,$ox,$num,$col){ $x=$gx+$c*$cw+($cw-4)/2-11+$ox; $y=$gy+$r*$ch+($ch-4)/2-11; Oval $x $y 22 22 $col | Out-Null; Txt $x $y 22 22 ([string]$num) 11 -1 $WHITE "Aptos" 2 3 | Out-Null }
DotAt 0 2 0 1 $RED      # scope creep (HH)
DotAt 1 1 -14 2 $AMBER  # vendor ambiguity (MM)
DotAt 1 2 0 3 $RED      # data quality (MH)
DotAt 2 2 0 4 $AMBER    # security gaps (LH)
DotAt 1 1 14 5 $AMBER   # adoption (MM)

# legend right
$risks = @(
 @{n=1;t="Scope creep"; m="Lock Phase 2 to BOM/effectivity/baseline/change; defer MES/APM."; c=$RED},
 @{n=2;t="Vendor ambiguity"; m="Vendor-neutral adapters; no single-backbone dependency."; c=$AMBER},
 @{n=3;t="Data quality"; m="Gate every release with seeded E2E + quality-issue promotion."; c=$RED},
 @{n=4;t="Security gaps"; m="OIDC/JWT, role-based authority and audit trail (mandatory)."; c=$AMBER},
 @{n=5;t="Adoption resistance"; m="Dashboard + shipyard scenarios to show value early."; c=$AMBER}
)
Txt 450 116 480 16 "Risk -> mitigation" 11 -1 $WHITE "Aptos" 1 1 | Out-Null
$ly=140
foreach($rk in $risks){ Panel 450 $ly 486 62 $PANEL $rk.c 0.75 0 0.06 | Out-Null; Oval 462 ($ly+10) 22 22 $rk.c | Out-Null; Txt 462 ($ly+10) 22 22 ([string]$rk.n) 11 -1 $WHITE "Aptos" 2 3 | Out-Null; Txt 494 ($ly+8) 430 16 $rk.t 10 -1 $WHITE "Aptos" 1 1 | Out-Null; Txt 494 ($ly+28) 432 28 $rk.m 8.5 0 $LIGHT "Aptos" 1 1 | Out-Null; $ly += 70 }

$pres.Save()
$s.Export("$outDir\s21_matrix.png","PNG",1280,720)
$pres.Close()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($pres) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($ppt)  | Out-Null
[GC]::Collect(); [GC]::WaitForPendingFinalizers()
Write-Output "BUILD8_S21_DONE"
