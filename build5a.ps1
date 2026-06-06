$ErrorActionPreference = "Stop"
$v2 = "D:\이력서\AVEVA - Marine Principal Technical Support & Consultant – PLM SME, Busan\Proposal\Future_Industrial_PLM_Meeting_Deck_EN_Rebuilt_v2.pptx"
$name = "Future_Industrial_PLM_Meeting_Deck_EN_Rebuilt_v2.pptx"
$outDir = "C:\Users\namma\.claude\plm_slide_work\review_png5"
New-Item -ItemType Directory -Force -Path $outDir | Out-Null

function Rgb([string]$hex){ $hex=$hex.TrimStart('#'); [int]([Convert]::ToInt32($hex.Substring(0,2),16)+[Convert]::ToInt32($hex.Substring(2,2),16)*256+[Convert]::ToInt32($hex.Substring(4,2),16)*65536) }
function Retry($sb){ for($a=0;$a -lt 6;$a++){ try{ return (& $sb) }catch{ Start-Sleep -Milliseconds 80 } }; return (& $sb) }

$WHITE="FFFFFF"; $LIGHT="C7D6E5"; $MUTE="93A8C0"
$CYAN="34C0EE"; $BLUE="4F8BE8"; $GREEN="49C28C"; $ORANGE="F0832F"; $VIOLET="9B7BE0"; $MAG="D957B8"
$BGD="0B1626"; $PANEL="14253C"; $BARC="12273C"; $HEXBG="241038"
$TAGA_BG="0E2E3A"; $TAGA_TX="56C8E8"; $TAGP_BG="231A3D"; $TAGP_TX="B79BE6"

$ppt = New-Object -ComObject PowerPoint.Application
foreach($p in @($ppt.Presentations)){ if($p.Name -eq $name){ try{ $p.Saved=$true; $p.Close() }catch{} } }
$pres=$ppt.Presentations.Open($v2,0,0,0)

function NoShadow($s){ try{$s.Shadow.Visible=0}catch{} }
function Panel($l,$t,$w,$h,$fill,$line,$lw,$trans,$rad){ $s=$script:slide.Shapes.AddShape(5,[single]$l,[single]$t,[single]$w,[single]$h); try{$s.Adjustments.Item(1)=[single]$rad}catch{}; $s.Fill.Solid(); Retry{$s.Fill.ForeColor.RGB=[int](Rgb $fill)}; $s.Fill.Transparency=[single]$trans; if($line){$s.Line.Visible=-1; Retry{$s.Line.ForeColor.RGB=[int](Rgb $line)}; $s.Line.Weight=[single]$lw}else{$s.Line.Visible=0}; NoShadow $s; return $s }
function Bar($l,$t,$w,$h,$fill,$rad){ $s=$script:slide.Shapes.AddShape(5,[single]$l,[single]$t,[single]$w,[single]$h); try{$s.Adjustments.Item(1)=[single]$rad}catch{}; $s.Fill.Solid(); Retry{$s.Fill.ForeColor.RGB=[int](Rgb $fill)}; $s.Line.Visible=0; NoShadow $s; return $s }
function Oval($l,$t,$w,$h,$fill){ $s=$script:slide.Shapes.AddShape(9,[single]$l,[single]$t,[single]$w,[single]$h); $s.Fill.Solid(); Retry{$s.Fill.ForeColor.RGB=[int](Rgb $fill)}; $s.Line.Visible=0; NoShadow $s; return $s }
function Txt($l,$t,$w,$h,$text,$size,$bold,$color,$font,$align,$anchor){ $tb=$script:slide.Shapes.AddTextbox(1,[single]$l,[single]$t,[single]$w,[single]$h); $tb.Fill.Visible=0;$tb.Line.Visible=0;NoShadow $tb; $tf=$tb.TextFrame;$tf.WordWrap=-1;$tf.AutoSize=0;$tf.MarginLeft=[single]2;$tf.MarginRight=[single]2;$tf.MarginTop=[single]1;$tf.MarginBottom=[single]1;$tf.VerticalAnchor=[int]$anchor; $tr=$tf.TextRange;$tr.Text=$text;$tr.Font.Size=[single]$size;$tr.Font.Bold=[int]$bold;$tr.Font.Name=[string]$font;Retry{$tr.Font.Color.RGB=[int](Rgb $color)};$tr.ParagraphFormat.Alignment=[int]$align; try{$tr.ParagraphFormat.SpaceBefore=[single]0;$tr.ParagraphFormat.SpaceAfter=[single]0}catch{}; return $tb }
function Chip($l,$t,$w,$h,$text,$size,$fill,$tcolor,$line,$font){ $s=$script:slide.Shapes.AddShape(5,[single]$l,[single]$t,[single]$w,[single]$h); try{$s.Adjustments.Item(1)=[single]0.4}catch{}; $s.Fill.Solid();Retry{$s.Fill.ForeColor.RGB=[int](Rgb $fill)}; if($line){$s.Line.Visible=-1;Retry{$s.Line.ForeColor.RGB=[int](Rgb $line)};$s.Line.Weight=[single]0.75}else{$s.Line.Visible=0}; NoShadow $s; $tf=$s.TextFrame;$tf.WordWrap=-1;$tf.AutoSize=0;$tf.VerticalAnchor=[int]3;$tf.MarginLeft=[single]2;$tf.MarginRight=[single]2;$tf.MarginTop=[single]0;$tf.MarginBottom=[single]0; $tr=$tf.TextRange;$tr.Text=$text;$tr.Font.Size=[single]$size;$tr.Font.Name=[string]$font;Retry{$tr.Font.Color.RGB=[int](Rgb $tcolor)};$tr.ParagraphFormat.Alignment=[int]2; return $s }
function Arrow($x1,$y1,$x2,$y2,$color,$w){ $c=$script:slide.Shapes.AddConnector(1,[single]$x1,[single]$y1,[single]$x2,[single]$y2); Retry{$c.Line.ForeColor.RGB=[int](Rgb $color)};$c.Line.Weight=[single]$w;$c.Line.BeginArrowheadStyle=[int]1;$c.Line.EndArrowheadStyle=[int]2;NoShadow $c; return $c }
function Hex($cx,$cy,$w,$h,$fill,$line,$text,$tcolor,$size){ $s=$script:slide.Shapes.AddShape(10,[single]($cx-$w/2),[single]($cy-$h/2),[single]$w,[single]$h); $s.Fill.Solid();Retry{$s.Fill.ForeColor.RGB=[int](Rgb $fill)};$s.Line.Visible=-1;Retry{$s.Line.ForeColor.RGB=[int](Rgb $line)};$s.Line.Weight=[single]1.25;NoShadow $s; $tf=$s.TextFrame;$tf.WordWrap=-1;$tf.AutoSize=0;$tf.VerticalAnchor=[int]3;$tf.MarginLeft=[single]1;$tf.MarginRight=[single]1;$tf.MarginTop=[single]0;$tf.MarginBottom=[single]0; $tr=$tf.TextRange;$tr.Text=$text;$tr.Font.Size=[single]$size;$tr.Font.Bold=[int](-1);$tr.Font.Name="Aptos";Retry{$tr.Font.Color.RGB=[int](Rgb $tcolor)};$tr.ParagraphFormat.Alignment=[int]2; return $s }
function DelAll($sl){ foreach($sh in @($sl.Shapes)){ try{$sh.Delete()}catch{} } }
function FindLike($pres,$prefix){ foreach($sl in $pres.Slides){ foreach($sh in $sl.Shapes){ try{ if($sh.HasTextFrame -eq -1 -and $sh.TextFrame.HasText -eq -1){ if($sh.TextFrame.TextRange.Text.Trim().StartsWith($prefix)){ return $sl } } }catch{} } }; return $null }

$flow = FindLike $pres "AI-gated"
if($flow -eq $null){ Write-Output "FLOW SLIDE NOT FOUND"; $pres.Close(); exit }
$flowIdx = [int]$flow.SlideIndex
Write-Output ("FLOW idx=" + $flowIdx)
DelAll $flow
$script:slide = $flow

# background + header
Bar 0 0 960 540 $BGD 0 | Out-Null
Txt 26 10 760 28 "AI-gated flow across AVEVA design tools and the API-built PLM" 20 -1 $WHITE "Aptos Display" 1 1 | Out-Null
Txt 770 12 158 14 "Future Industrial PLM" 8.5 0 $MUTE "Aptos" 3 1 | Out-Null
Txt 905 28 30 14 ([string]$flowIdx) 9 0 $MUTE "Aptos" 3 1 | Out-Null
Txt 26 40 908 16 "AVEVA E3D/Marine 3D  ->  AVEVA Draw 2D (AutoCAD DWG/DXF)  ->  MTO/E-BOM  ->  approval  ->  effectivity  ->  baseline  ->  ECR/ECO  ->  M-BOM/MCO  ->  impact" 9.5 0 $LIGHT "Aptos" 1 1 | Out-Null

# thin PLM status bar
Bar 22 64 916 22 $BARC 0.25 | Out-Null
Txt 34 64 600 22 "Future PLM   -   API workspace   ·   served at /api/v1   ·   live status tracked" 9 0 $MUTE "Consolas" 1 3 | Out-Null
Oval 854 71 8 8 $GREEN | Out-Null
Txt 866 64 70 22 "RUNNING" 8.5 -1 $GREEN "Aptos" 1 3 | Out-Null
Txt 36 90 880 16 "Design in AVEVA (E3D/Marine + Draw)  ->  governed & tracked in the API-built PLM, AI-gated at every step" 10.5 -1 $CYAN "Aptos" 1 1 | Out-Null

# 9 cards (serpentine 3x3)
$cards = @(
 @{x=40;  y=116; n=1; t="AVEVA E3D Design";  tag="AVEVA E3D/Marine"; ta=1; d="Hull & outfitting 3D model"; st="modeled"; c=$CYAN},
 @{x=360; y=116; n=2; t="2D Production Drawing"; tag="AVEVA Draw"; ta=1; d="Generate 2D sheet -> export AutoCAD DWG / DXF"; st="DWG issued"; c=$CYAN},
 @{x=680; y=116; n=3; t="E-BOM / MTO";        tag="AVEVA MTO"; ta=1; d="Material take-off / eng. BOM from model"; st="extracted"; c=$CYAN},
 @{x=680; y=250; n=4; t="Approval";           tag="PLM · API"; ta=0; d="Approver signs off in the API-built PLM"; st="approved"; c=$BLUE},
 @{x=360; y=250; n=5; t="Effectivity";        tag="PLM · API"; ta=0; d="Hull / Block / Zone / Date"; st="scoped"; c=$GREEN},
 @{x=40;  y=250; n=6; t="Baseline";           tag="PLM · API"; ta=0; d="Frozen configuration / released rev"; st="baselined"; c=$GREEN},
 @{x=40;  y=384; n=7; t="ECR -> ECO";         tag="PLM · API"; ta=0; d="Change request -> change order"; st="ECO closed"; c=$ORANGE},
 @{x=360; y=384; n=8; t="M-BOM + MCO";        tag="PLM·API + ERM"; ta=0; d="Manufacturing BOM + mfg change order"; st="released"; c=$ORANGE},
 @{x=680; y=384; n=9; t="Impact + Approval";  tag="PLM · API"; ta=0; d="Affected objects + approval"; st="approved"; c=$VIOLET}
)
foreach($k in $cards){
    $x=$k.x; $y=$k.y
    Panel $x $y 240 94 $PANEL $k.c 1 0 0.07 | Out-Null
    Bar $x $y 240 4 $k.c 0 | Out-Null
    Oval ($x+8) ($y+8) 18 18 $k.c | Out-Null
    Txt ($x+8) ($y+8) 18 18 ([string]$k.n) 10 -1 $WHITE "Aptos" 2 3 | Out-Null
    Txt ($x+30) ($y+6) 126 16 $k.t 9.5 -1 $WHITE "Aptos" 1 1 | Out-Null
    if($k.ta -eq 1){ Chip ($x+158) ($y+8) 76 13 $k.tag 7 $TAGA_BG $TAGA_TX $TAGA_TX "Aptos" | Out-Null } else { Chip ($x+158) ($y+8) 76 13 $k.tag 7 $TAGP_BG $TAGP_TX $TAGP_TX "Aptos" | Out-Null }
    Txt ($x+10) ($y+28) 222 34 $k.d 8 0 $LIGHT "Aptos" 1 1 | Out-Null
    Chip ($x+10) ($y+68) 150 18 ($k.st + "  OK") 8 "15402F" $GREEN $GREEN "Aptos" | Out-Null
}

# arrows
Arrow 280 163 360 163 $MUTE 1.75 | Out-Null
Arrow 600 163 680 163 $MUTE 1.75 | Out-Null
Arrow 800 210 800 250 $MUTE 1.75 | Out-Null
Arrow 680 297 600 297 $MUTE 1.75 | Out-Null
Arrow 360 297 280 297 $MUTE 1.75 | Out-Null
Arrow 160 344 160 384 $MUTE 1.75 | Out-Null
Arrow 280 431 360 431 $MUTE 1.75 | Out-Null
Arrow 600 431 680 431 $MUTE 1.75 | Out-Null

# AI hexes
$hexPts = @(@(320,163),@(640,163),@(800,230),@(640,297),@(320,297),@(160,364),@(320,431),@(640,431))
foreach($p in $hexPts){ Hex $p[0] $p[1] 46 40 $HEXBG $MAG "AI" $MAG 9.5 | Out-Null }

# legend + status strip
Txt 36 486 880 14 "AI gate at every step = error check + anomaly diagnosis  ·  PASS -> auto-promote  ·  FAIL -> block & flag (no promote)" 8.5 0 $MAG "Aptos" 1 1 | Out-Null
Txt 36 506 880 16 "PLM live status:  modeled > DWG issued > MTO extracted > approved > scoped > baselined > ECO closed > released > approved" 7.5 -1 $GREEN "Aptos" 1 1 | Out-Null

$pres.Save()
$flow.Export("$outDir\flow.png","PNG",1280,720)
$pres.Close()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($pres) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($ppt)  | Out-Null
[GC]::Collect(); [GC]::WaitForPendingFinalizers()
Write-Output "BUILD5A_DONE"
