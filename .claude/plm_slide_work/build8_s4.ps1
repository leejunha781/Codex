$ErrorActionPreference = "Stop"
$v2 = "D:\이력서\AVEVA - Marine Principal Technical Support & Consultant – PLM SME, Busan\Proposal\Future_Industrial_PLM_Meeting_Deck_EN_Rebuilt_v2.pptx"
$name = "Future_Industrial_PLM_Meeting_Deck_EN_Rebuilt_v2.pptx"
$outDir = "C:\Users\namma\.claude\plm_slide_work\review_png8"

function Rgb([string]$hex){ $hex=$hex.TrimStart('#'); [int]([Convert]::ToInt32($hex.Substring(0,2),16)+[Convert]::ToInt32($hex.Substring(2,2),16)*256+[Convert]::ToInt32($hex.Substring(4,2),16)*65536) }
function Retry($sb){ for($a=0;$a -lt 6;$a++){ try{ return (& $sb) }catch{ Start-Sleep -Milliseconds 80 } }; return (& $sb) }
$WHITE="FFFFFF"; $LIGHT="C7D6E5"; $MUTE="93A8C0"
$CYAN="34C0EE"; $BLUE="4F8BE8"; $GREEN="49C28C"; $ORANGE="F0832F"; $VIOLET="9B7BE0"; $SLATE2="8FA9C4"
$BGD="0B1626"; $PLOT="0A1626"; $GRID="24405E"; $PANEL="14253C"

$ppt = New-Object -ComObject PowerPoint.Application
foreach($p in @($ppt.Presentations)){ if($p.Name -eq $name){ try{ $p.Saved=$true; $p.Close() }catch{} } }
$pres=$ppt.Presentations.Open($v2,0,0,0)

function NoShadow($s){ try{$s.Shadow.Visible=0}catch{} }
function Panel($l,$t,$w,$h,$fill,$line,$lw,$trans,$rad){ $s=$script:slide.Shapes.AddShape(5,[single]$l,[single]$t,[single]$w,[single]$h); try{$s.Adjustments.Item(1)=[single]$rad}catch{}; $s.Fill.Solid(); Retry{$s.Fill.ForeColor.RGB=[int](Rgb $fill)}; $s.Fill.Transparency=[single]$trans; if($line){$s.Line.Visible=-1; Retry{$s.Line.ForeColor.RGB=[int](Rgb $line)}; $s.Line.Weight=[single]$lw}else{$s.Line.Visible=0}; NoShadow $s; return $s }
function Oval($l,$t,$w,$h,$fill){ $s=$script:slide.Shapes.AddShape(9,[single]$l,[single]$t,[single]$w,[single]$h); $s.Fill.Solid(); Retry{$s.Fill.ForeColor.RGB=[int](Rgb $fill)}; $s.Line.Visible=0; NoShadow $s; return $s }
function EllipseT($l,$t,$w,$h,$fill,$trans,$line){ $s=$script:slide.Shapes.AddShape(9,[single]$l,[single]$t,[single]$w,[single]$h); $s.Fill.Solid(); Retry{$s.Fill.ForeColor.RGB=[int](Rgb $fill)}; $s.Fill.Transparency=[single]$trans; $s.Line.Visible=-1; Retry{$s.Line.ForeColor.RGB=[int](Rgb $line)}; $s.Line.Weight=[single]1; try{$s.Line.DashStyle=4}catch{}; NoShadow $s; return $s }
function Txt($l,$t,$w,$h,$text,$size,$bold,$color,$font,$align,$anchor){ $tb=$script:slide.Shapes.AddTextbox(1,[single]$l,[single]$t,[single]$w,[single]$h); $tb.Fill.Visible=0;$tb.Line.Visible=0;NoShadow $tb; $tf=$tb.TextFrame;$tf.WordWrap=-1;$tf.AutoSize=0;$tf.MarginLeft=[single]2;$tf.MarginRight=[single]2;$tf.MarginTop=[single]1;$tf.MarginBottom=[single]1;$tf.VerticalAnchor=[int]$anchor; $tr=$tf.TextRange;$tr.Text=$text;$tr.Font.Size=[single]$size;$tr.Font.Bold=[int]$bold;$tr.Font.Name=[string]$font;Retry{$tr.Font.Color.RGB=[int](Rgb $color)};$tr.ParagraphFormat.Alignment=[int]$align; try{$tr.ParagraphFormat.SpaceBefore=[single]0;$tr.ParagraphFormat.SpaceAfter=[single]0}catch{}; return $tb }
function Arrow($x1,$y1,$x2,$y2,$color,$w){ $c=$script:slide.Shapes.AddConnector(1,[single]$x1,[single]$y1,[single]$x2,[single]$y2); Retry{$c.Line.ForeColor.RGB=[int](Rgb $color)};$c.Line.Weight=[single]$w;$c.Line.BeginArrowheadStyle=[int]1;$c.Line.EndArrowheadStyle=[int]2;NoShadow $c; return $c }
function DelAll($sl){ foreach($sh in @($sl.Shapes)){ try{$sh.Delete()}catch{} } }
function FindLike($pres,$prefix){ foreach($sl in $pres.Slides){ foreach($sh in $sl.Shapes){ try{ if($sh.HasTextFrame -eq -1 -and $sh.TextFrame.HasText -eq -1){ if($sh.TextFrame.TextRange.Text.Trim().StartsWith($prefix)){ return $sl } } }catch{} } }; return $null }

$s4 = FindLike $pres "Current PLM landscape"
Write-Output ("s4 idx=" + $s4.SlideIndex)
DelAll $s4
$script:slide = $s4

$bg=$script:slide.Shapes.AddShape(1,[single]0,[single]0,[single]960,[single]540); $bg.Line.Visible=0; NoShadow $bg; $bg.Fill.Solid(); Retry{$bg.Fill.ForeColor.RGB=[int](Rgb $BGD)}
Txt 26 14 760 30 "Current PLM landscape" 26 -1 $WHITE "Aptos Display" 1 1 | Out-Null
Txt 770 16 158 14 ([string]$s4.SlideIndex) 9 0 $MUTE "Aptos" 3 1 | Out-Null
Txt 26 50 760 16 "Where AVEVA differentiates: bridging engineering and operations" 11 0 $LIGHT "Aptos" 1 1 | Out-Null

# plot frame
Panel 56 86 556 388 $PLOT $GRID 1 0 0.02 | Out-Null
# axes
Arrow 80 300 596 300 $GRID 1.5 | Out-Null
Arrow 334 466 334 96 $GRID 1.5 | Out-Null
Txt 470 302 138 14 "Operations / Asset >" 7.5 0 $MUTE "Aptos" 3 1 | Out-Null
Txt 66 302 150 14 "< Engineering / PLM" 7.5 0 $MUTE "Aptos" 1 1 | Out-Null
Txt 340 100 200 12 "Unified platform (up)" 7.5 0 $MUTE "Aptos" 1 1 | Out-Null
Txt 340 452 200 12 "Siloed point tools (down)" 7.5 0 $MUTE "Aptos" 1 1 | Out-Null
# sweet spot
EllipseT 234 120 248 100 $CYAN 0.84 $CYAN | Out-Null
Txt 250 124 220 12 "Engineering -> Operations sweet spot" 7.5 -1 $CYAN "Aptos" 1 1 | Out-Null
# bubbles
Oval 214 213 12 12 $SLATE2 | Out-Null
Txt 230 211 80 14 "Siemens" 8 -1 $LIGHT "Aptos" 1 1 | Out-Null
Oval 256 185 12 12 $VIOLET | Out-Null
Txt 272 183 80 14 "Dassault" 8 -1 $LIGHT "Aptos" 1 1 | Out-Null
Oval 352 283 12 12 $ORANGE | Out-Null
Txt 368 281 60 14 "PTC" 8 -1 $LIGHT "Aptos" 1 1 | Out-Null
Oval 478 233 12 12 $GREEN | Out-Null
Txt 494 231 80 14 "Hexagon" 8 -1 $LIGHT "Aptos" 1 1 | Out-Null
Oval 344 160 20 20 $CYAN | Out-Null
Txt 368 160 90 18 "AVEVA" 11 -1 $CYAN "Aptos" 1 1 | Out-Null

# right key
Txt 632 90 300 16 "Stack & position" 10 -1 $WHITE "Aptos" 1 1 | Out-Null
$key = @(
 @{c=$CYAN;   v="AVEVA";    s="E3D/Marine · AIM · PI · CONNECT"; p="Engineering -> Operations (the bridge)"},
 @{c=$SLATE2; v="Siemens";  s="Teamcenter · NX · Opcenter"; p="PLM core backbone"},
 @{c=$VIOLET; v="Dassault"; s="3DEXPERIENCE · CATIA · DELMIA"; p="Experience / virtual-twin platform"},
 @{c=$GREEN;  v="Hexagon";  s="SDx · Smart 3D · j5 · Completions"; p="Asset lifecycle & operations"},
 @{c=$ORANGE; v="PTC";      s="Windchill · Codebeamer · ThingWorx"; p="PLM + IoT digital thread"}
)
$ky=116
foreach($k in $key){
  Oval 636 ($ky+3) 12 12 $k.c | Out-Null
  Txt 656 ($ky) 270 15 $k.v 9.5 -1 $WHITE "Aptos" 1 1 | Out-Null
  Txt 656 ($ky+16) 280 13 $k.s 7.5 0 $LIGHT "Consolas" 1 1 | Out-Null
  Txt 656 ($ky+30) 280 13 $k.p 7.5 0 $MUTE "Aptos" 1 1 | Out-Null
  $ky += 54
}
# takeaway
Panel 56 484 878 34 "0E2A24" $GREEN 1 0 0.1 | Out-Null
Txt 66 484 858 34 "Takeaway: don't compete as just another PLM backbone — build the engineering-to-operations layer that existing PLM, CAD, MES and historian platforms cannot easily unify." 9 -1 $WHITE "Aptos" 1 3 | Out-Null

$pres.Save()
$s4.Export("$outDir\s4_map.png","PNG",1280,720)
$pres.Close()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($pres) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($ppt)  | Out-Null
[GC]::Collect(); [GC]::WaitForPendingFinalizers()
Write-Output "BUILD8_S4_DONE"
