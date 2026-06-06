$ErrorActionPreference = "Stop"
$v2 = "D:\이력서\AVEVA - Marine Principal Technical Support & Consultant – PLM SME, Busan\Proposal\Future_Industrial_PLM_Meeting_Deck_EN_Rebuilt_v2.pptx"
$name = "Future_Industrial_PLM_Meeting_Deck_EN_Rebuilt_v2.pptx"
$outDir = "C:\Users\namma\.claude\plm_slide_work\review_png2"
New-Item -ItemType Directory -Force -Path $outDir | Out-Null

function Rgb([string]$hex){ $hex=$hex.TrimStart('#'); [int]([Convert]::ToInt32($hex.Substring(0,2),16) + [Convert]::ToInt32($hex.Substring(2,2),16)*256 + [Convert]::ToInt32($hex.Substring(4,2),16)*65536) }

$WHITE="FFFFFF"; $LIGHT="C7D6E5"; $MUTE="93A8C0"
$CYAN="34C0EE"; $BLUE="4F8BE8"; $GREEN="49C28C"; $ORANGE="F0832F"; $VIOLET="9B7BE0"
$GOLD="E8C24A"; $AMBER="E8B23A"; $RED="E5634E"
$PANEL="14253C"; $PANEL_LINE="2C4663"; $TRACK="0D1B2D"; $GRID="24405E"

$ppt = New-Object -ComObject PowerPoint.Application
$pres = $null
foreach($p in $ppt.Presentations){ if($p.Name -eq $name){ $pres=$p; break } }
$openedByUs = $false
if($pres -eq $null){ $pres = $ppt.Presentations.Open($v2, 0, 0, 0); $openedByUs = $true }

# ---------- helpers (operate on $script:slide) ----------
function NoShadow($s){ try{$s.Shadow.Visible=0}catch{} }
function Panel($l,$t,$w,$h,$fill,$line,$lw,$trans,$rad){
    $s=$script:slide.Shapes.AddShape(5,[single]$l,[single]$t,[single]$w,[single]$h)
    try{$s.Adjustments.Item(1)=[single]$rad}catch{}
    $s.Fill.Solid(); $s.Fill.ForeColor.RGB=[int](Rgb $fill); $s.Fill.Transparency=[single]$trans
    if($line){ $s.Line.Visible=-1; $s.Line.ForeColor.RGB=[int](Rgb $line); $s.Line.Weight=[single]$lw } else { $s.Line.Visible=0 }
    NoShadow $s; return $s
}
function Bar($l,$t,$w,$h,$fill,$rad){
    $s=$script:slide.Shapes.AddShape(5,[single]$l,[single]$t,[single]$w,[single]$h)
    try{$s.Adjustments.Item(1)=[single]$rad}catch{}
    $s.Fill.Solid(); $s.Fill.ForeColor.RGB=[int](Rgb $fill); $s.Line.Visible=0; NoShadow $s; return $s
}
function Txt($l,$t,$w,$h,$text,$size,$bold,$color,$font,$align,$anchor){
    $tb=$script:slide.Shapes.AddTextbox(1,[single]$l,[single]$t,[single]$w,[single]$h)
    $tb.Fill.Visible=0; $tb.Line.Visible=0; NoShadow $tb
    $tf=$tb.TextFrame; $tf.WordWrap=-1; $tf.AutoSize=0
    $tf.MarginLeft=[single]2;$tf.MarginRight=[single]2;$tf.MarginTop=[single]1;$tf.MarginBottom=[single]1
    $tf.VerticalAnchor=[int]$anchor
    $tr=$tf.TextRange; $tr.Text=$text
    $tr.Font.Size=[single]$size; $tr.Font.Bold=[int]$bold; $tr.Font.Name=[string]$font; $tr.Font.Color.RGB=[int](Rgb $color)
    $tr.ParagraphFormat.Alignment=[int]$align
    try{$tr.ParagraphFormat.SpaceBefore=[single]0;$tr.ParagraphFormat.SpaceAfter=[single]0}catch{}
    return $tb
}
function Chip($l,$t,$w,$h,$text,$size,$fill,$tcolor,$line,$font){
    $s=$script:slide.Shapes.AddShape(5,[single]$l,[single]$t,[single]$w,[single]$h)
    try{$s.Adjustments.Item(1)=[single]0.35}catch{}
    $s.Fill.Solid(); $s.Fill.ForeColor.RGB=[int](Rgb $fill)
    if($line){ $s.Line.Visible=-1; $s.Line.ForeColor.RGB=[int](Rgb $line); $s.Line.Weight=[single]0.75 } else { $s.Line.Visible=0 }
    NoShadow $s
    $tf=$s.TextFrame; $tf.WordWrap=-1; $tf.AutoSize=0; $tf.VerticalAnchor=[int]3
    $tf.MarginLeft=[single]2;$tf.MarginRight=[single]2;$tf.MarginTop=[single]0;$tf.MarginBottom=[single]0
    $tr=$tf.TextRange; $tr.Text=$text; $tr.Font.Size=[single]$size; $tr.Font.Name=[string]$font
    $tr.Font.Color.RGB=[int](Rgb $tcolor); $tr.ParagraphFormat.Alignment=[int]2
    return $s
}
function Line2($x1,$y1,$x2,$y2,$color,$w){
    $c=$script:slide.Shapes.AddConnector(1,[single]$x1,[single]$y1,[single]$x2,[single]$y2)
    $c.Line.ForeColor.RGB=[int](Rgb $color); $c.Line.Weight=[single]$w
    $c.Line.BeginArrowheadStyle=[int]1; $c.Line.EndArrowheadStyle=[int]1; NoShadow $c; return $c
}
function Diamond($cx,$cy,$r,$fill){
    $s=$script:slide.Shapes.AddShape(4,[single]($cx-$r),[single]($cy-$r),[single]($r*2),[single]($r*2)) # 4=diamond
    $s.Fill.Solid(); $s.Fill.ForeColor.RGB=[int](Rgb $fill); $s.Line.Visible=0; NoShadow $s; return $s
}
function DelShapes($sl,$names){ foreach($sh in @($sl.Shapes)){ if($names -contains $sh.Name){ $sh.Delete() } } }
function SetFirstChar($sl,$shapeName,$ch){ foreach($sh in $sl.Shapes){ if($sh.Name -eq $shapeName){ try{$sh.TextFrame.TextRange.Characters(1,1).Text=$ch}catch{} } } }

# ================= SLIDE 10: relabel B,C,D,E -> A,B,C,D =================
$s10 = $pres.Slides.Item(10)
SetFirstChar $s10 'Text 5' 'A'
SetFirstChar $s10 'Text 6' 'B'
SetFirstChar $s10 'Text 7' 'C'
SetFirstChar $s10 'Text 8' 'D'

# ================= SLIDE 13: WBS timeline =================
$s13 = $pres.Slides.Item(13)
DelShapes $s13 @('Text 4','Text 5','Text 6','Text 7','Text 8')
$script:slide = $s13

$x0 = 188.0; $perWk = 35.9
function WX([double]$w){ 188.0 + $perWk*$w }
$rowTops = @(150,206,262,318,374)
$wbs = @(
 @{n="WBS 0"; nm="Readiness";          s=0;  e=2;  c=$CYAN;   d="2 wks"; sc="Backlog · env · source · roles"},
 @{n="WBS 1"; nm="Phase 1 hardening";  s=2;  e=6;  c=$BLUE;   d="4 wks"; sc="S6 gates · OIDC · dashboard · CI"},
 @{n="WBS 2"; nm="Configuration core"; s=6;  e=12; c=$GREEN;  d="6 wks"; sc="BOM · effectivity · baseline · APIs"},
 @{n="WBS 3"; nm="Change impact";      s=12; e=16; c=$ORANGE; d="4 wks"; sc="ECR/ECO/MOC · rules · promote · UI"},
 @{n="WBS 4"; nm="Pilot adoption";     s=16; e=20; c=$VIOLET; d="4 wks"; sc="Shipyard UAT · perf · security"}
)
# gate boundaries + gridlines + axis labels
$gates = @(2,6,12,16,20)
Line2 (WX 0) 145 (WX 0) 420 $GRID 0.75 | Out-Null
Txt ((WX 0)-26) 104 52 13 "start" 8 0 $MUTE "Aptos" 2 1 | Out-Null
$gi=1
foreach($g in $gates){
    $gx = WX $g
    Line2 $gx 140 $gx 420 $GRID 0.75 | Out-Null
    Diamond $gx 132 5 $GOLD | Out-Null
    Txt ($gx-30) 90 60 13 ("G"+$gi) 9 -1 $GOLD "Aptos" 2 1 | Out-Null
    Txt ($gx-30) 104 60 13 ("wk"+$g) 8 0 $MUTE "Aptos" 2 1 | Out-Null
    $gi++
}
Txt 36 124 880 13 "Sequential delivery, 20 weeks. Each package closes at a KPI acceptance gate (G1-G5)." 8.5 0 $LIGHT "Aptos" 1 1 | Out-Null

# bars + labels
for($i=0;$i -lt 5;$i++){
    $w=$wbs[$i]; $rt=$rowTops[$i]
    $bx = WX $w.s; $bw = (WX $w.e) - $bx
    # left label
    Txt 36 ($rt+1) 148 16 $w.n 11 -1 $w.c "Aptos" 1 1 | Out-Null
    Txt 36 ($rt+18) 148 14 $w.nm 9 0 $WHITE "Aptos" 1 1 | Out-Null
    # bar
    Bar $bx ($rt+4) $bw 24 $w.c 0.4 | Out-Null
    Txt $bx ($rt+4) $bw 24 $w.d 10 -1 $WHITE "Aptos" 2 3 | Out-Null
    # scope under bar
    $scW = [math]::Min(262, 914 - $bx)
    Txt $bx ($rt+30) $scW 13 $w.sc 7.5 0 $MUTE "Aptos" 1 1 | Out-Null
}
# WBS -> KPI gate mapping (one line)
Txt 36 430 880 14 "Gates -> KPI: G1 env & backlog ready  |  G2 E2E >=95% / OIDC  |  G3 BOM & Effectivity 99%  |  G4 Impact precision >=85%  |  G5 Thread >=85% / dashboard <=5s" 8 0 $GOLD "Aptos" 1 1 | Out-Null

# ================= SLIDE 14: KPI scorecard 4x2 =================
$s14 = $pres.Slides.Item(14)
DelShapes $s14 @('Text 4','Text 5','Text 6','Text 7','Text 8','Text 9','Text 10','Text 11')
$script:slide = $s14

$colX = @(36,262,488,714); $rowY = @(100,300); $cardW=210; $cardH=180
$kpis = @(
 @{t="API contract coverage"; v="100%";  p=100; c=$CYAN;   note="All committed features ship with OpenAPI + tests"},
 @{t="E2E gate pass rate";     v=">=95%"; p=95;  c=$CYAN;   note="CI validates schema, seed, evaluate and promote"},
 @{t="Digital thread coverage";v=">=85%"; p=85;  c=$GREEN;  note="Equipment linked to tag, document, 3D, asset, test"},
 @{t="BOM expansion accuracy"; v=">=99%"; p=99;  c=$GREEN;  note="Recursive BOM returns correct levels & quantities"},
 @{t="Effectivity accuracy";   v=">=99%"; p=99;  c=$GREEN;  note="Hull / block / option / date filters correct"},
 @{t="Impact candidate precision"; v=">=85%"; p=85; c=$GREEN; note="Rule engine suggests relevant affected objects"},
 @{t="Dashboard response";     v="<=5 s"; p=55;  c=$ORANGE; note="KPI & readiness summary under target load"}
)
for($i=0;$i -lt 7;$i++){
    $k=$kpis[$i]
    $cx=$colX[$i % 4]; $cy=$rowY[[math]::Floor($i/4)]
    Panel $cx $cy $cardW $cardH $PANEL $PANEL_LINE 1 0 0.06 | Out-Null
    Bar $cx $cy $cardW 5 $k.c 0 | Out-Null
    Txt ($cx+14) ($cy+12) ($cardW-28) 34 $k.t 11 -1 $WHITE "Aptos" 1 1 | Out-Null
    Txt $cx ($cy+48) $cardW 42 $k.v 30 -1 $k.c "Aptos Display" 2 1 | Out-Null
    Bar ($cx+16) ($cy+106) ($cardW-32) 9 $TRACK 0.5 | Out-Null
    Bar ($cx+16) ($cy+106) ([double]($cardW-32)*$k.p/100.0) 9 $k.c 0.5 | Out-Null
    Txt ($cx+14) ($cy+122) ($cardW-28) 46 $k.note 8 0 $LIGHT "Aptos" 1 1 | Out-Null
}
# 8th card: gate decision
$cx=$colX[3]; $cy=$rowY[1]
Panel $cx $cy $cardW $cardH $PANEL $PANEL_LINE 1 0 0.06 | Out-Null
Bar $cx $cy $cardW 5 $VIOLET 0 | Out-Null
Txt ($cx+14) ($cy+12) ($cardW-28) 18 "Gate decision" 11 -1 $WHITE "Aptos" 1 1 | Out-Null
Chip ($cx+16) ($cy+44) ($cardW-32) 28 "PASS" 11 "15402F" $GREEN $GREEN "Aptos" | Out-Null
Chip ($cx+16) ($cy+82) ($cardW-32) 28 "CONDITIONAL PASS + backlog" 8.5 "3A2E10" $AMBER $AMBER "Aptos" | Out-Null
Chip ($cx+16) ($cy+120) ($cardW-32) 28 "FAIL + remediation sprint" 8.5 "3A1614" $RED $RED "Aptos" | Out-Null

# ---------- save + export ----------
$pres.Save()
foreach($idx in @(10,13,14)){ $pres.Slides.Item($idx).Export("$outDir\slide-$idx.png","PNG",1280,720) }
if($openedByUs){ $pres.Close() }
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($pres) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($ppt)  | Out-Null
[GC]::Collect(); [GC]::WaitForPendingFinalizers()
Write-Output "BUILD2_DONE"
