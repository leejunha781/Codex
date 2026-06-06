$ErrorActionPreference = "Stop"
$v2 = "D:\이력서\AVEVA - Marine Principal Technical Support & Consultant – PLM SME, Busan\Proposal\Future_Industrial_PLM_Meeting_Deck_EN_Rebuilt_v2.pptx"
$name = "Future_Industrial_PLM_Meeting_Deck_EN_Rebuilt_v2.pptx"
$png = "C:\Users\namma\.claude\plm_slide_work\slide_flow.png"

function Rgb([string]$hex){ $hex=$hex.TrimStart('#'); [int]([Convert]::ToInt32($hex.Substring(0,2),16)+[Convert]::ToInt32($hex.Substring(2,2),16)*256+[Convert]::ToInt32($hex.Substring(4,2),16)*65536) }
function Retry($sb){ for($a=0;$a -lt 6;$a++){ try{ return (& $sb) }catch{ Start-Sleep -Milliseconds 80 } }; return (& $sb) }

$WHITE="FFFFFF"; $LIGHT="C7D6E5"; $MUTE="93A8C0"
$CYAN="34C0EE"; $BLUE="4F8BE8"; $GREEN="49C28C"; $ORANGE="F0832F"; $VIOLET="9B7BE0"; $MAG="D957B8"
$BGD="0B1626"; $PANEL="14253C"; $FRAME="0C1A2C"; $FRLINE="2C4663"; $BARC="15273C"; $HEXBG="241038"

$ppt = New-Object -ComObject PowerPoint.Application
# force-close any open/orphaned copy of v2 (discard unsaved), then open fresh from disk
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

# ---- append blank slide ----
$idx = [int]($pres.Slides.Count + 1)
$script:slide = $pres.Slides.Add($idx, 12)

# background
Bar 0 0 960 540 $BGD 0 | Out-Null
# title + subtitle
Txt 26 10 908 28 "AI-gated Engineering -> Manufacturing flow, served & tracked in the API-built PLM" 20 -1 $WHITE "Aptos Display" 1 1 | Out-Null
Txt 26 40 908 16 "E3D design -> E-BOM -> approval -> effectivity -> baseline -> ECR/ECO -> M-BOM/MCO -> impact.  AI checks every step; pass = promote, fail = flag." 10 0 $LIGHT "Aptos" 1 1 | Out-Null

# ---- PLM application window frame ----
Panel 22 64 916 452 $FRAME $FRLINE 1 0 0.02 | Out-Null
Bar 22 64 916 24 $BARC 0.02 | Out-Null
Oval 34 71 9 9 "FF5F56" | Out-Null
Oval 48 71 9 9 "FFBD2E" | Out-Null
Oval 62 71 9 9 "27C93F" | Out-Null
Txt 80 64 560 24 "Future PLM  -  API workspace  ·  served at /api/v1" 9 0 $MUTE "Consolas" 1 3 | Out-Null
Oval 820 72 8 8 $GREEN | Out-Null
Txt 832 64 100 24 "RUNNING" 8.5 -1 $GREEN "Aptos" 1 3 | Out-Null
Txt 36 96 880 16 "Engineering -> Manufacturing digital thread  —  API-served, AI-gated at every step" 11 -1 $CYAN "Aptos" 1 1 | Out-Null

# ---- stage cards (serpentine 2x4) ----
$cards = @(
 @{x=40;  y=124; n=1; t="AVEVA E3D Design";  d="3D engineering design model"; st="modeled"; c=$CYAN},
 @{x=280; y=124; n=2; t="E-BOM + Submit";    d="Engineering BOM created, submitted for approval"; st="submitted"; c=$CYAN},
 @{x=520; y=124; n=3; t="Approval in PLM";   d="Approver signs off in the API-built PLM"; st="approved"; c=$BLUE},
 @{x=760; y=124; n=4; t="Effectivity";       d="Hull / Block / Option / Date"; st="scoped"; c=$GREEN},
 @{x=760; y=300; n=5; t="Baseline";          d="Frozen configuration"; st="baselined"; c=$GREEN},
 @{x=520; y=300; n=6; t="ECR -> ECO";        d="Change request raised, change order executed"; st="ECO closed"; c=$ORANGE},
 @{x=280; y=300; n=7; t="M-BOM + MCO";       d="Manufacturing BOM + mfg change order"; st="released"; c=$ORANGE},
 @{x=40;  y=300; n=8; t="Impact + Approval"; d="Affected objects identified + approved"; st="approved"; c=$VIOLET}
)
foreach($k in $cards){
    $x=$k.x; $y=$k.y
    Panel $x $y 160 112 $PANEL $k.c 1 0 0.08 | Out-Null
    Bar $x $y 160 5 $k.c 0 | Out-Null
    Oval ($x+8) ($y+10) 22 22 $k.c | Out-Null
    Txt ($x+8) ($y+10) 22 22 ([string]$k.n) 11 -1 $WHITE "Aptos" 2 3 | Out-Null
    Txt ($x+34) ($y+8) 120 30 $k.t 9.5 -1 $WHITE "Aptos" 1 1 | Out-Null
    Txt ($x+10) ($y+42) 142 36 $k.d 8 0 $LIGHT "Aptos" 1 1 | Out-Null
    Chip ($x+10) ($y+86) 142 20 ($k.st + "  OK") 8 "15402F" $GREEN $GREEN "Aptos" | Out-Null
}

# ---- flow arrows ----
Arrow 200 180 280 180 $MUTE 1.75 | Out-Null
Arrow 440 180 520 180 $MUTE 1.75 | Out-Null
Arrow 680 180 760 180 $MUTE 1.75 | Out-Null
Arrow 840 236 840 300 $MUTE 1.75 | Out-Null
Arrow 760 356 680 356 $MUTE 1.75 | Out-Null
Arrow 520 356 440 356 $MUTE 1.75 | Out-Null
Arrow 280 356 200 356 $MUTE 1.75 | Out-Null

# ---- AI gate hexagons (on each transition) ----
$hexPts = @(@(240,180),@(480,180),@(720,180),@(840,268),@(720,356),@(480,356),@(240,356))
foreach($p in $hexPts){ Hex $p[0] $p[1] 48 42 $HEXBG $MAG "AI" $MAG 10 | Out-Null }

# ---- legend + caption + status strip ----
Txt 36 426 880 14 "AI gate = automated error check + anomaly diagnosis.   PASS -> auto-promote to next step.   FAIL -> block & flag for rework (no promote)." 8.5 0 $MAG "Aptos" 1 1 | Out-Null
Txt 36 446 880 28 "Every artifact (E-BOM, baseline, M-BOM, impact set) and every AI gate result is created and shown with live status inside the API-built PLM." 8.5 0 $LIGHT "Aptos" 1 1 | Out-Null
Txt 36 486 880 16 "PLM live status:  modeled > submitted > approved > scoped > baselined > ECO closed > released > approved" 8 -1 $GREEN "Aptos" 1 1 | Out-Null

# slide number
Txt 920 9 18 12 ([string]$idx) 9 0 $MUTE "Aptos" 3 1 | Out-Null

$pres.Save()
$pres.Slides.Item($idx).Export($png,"PNG",1280,720)
$pres.Close()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($pres) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($ppt)  | Out-Null
[GC]::Collect(); [GC]::WaitForPendingFinalizers()
Write-Output ("BUILD4_DONE new_slide=" + $idx)
