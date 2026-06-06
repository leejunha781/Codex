$ErrorActionPreference = "Stop"
$v2 = "D:\이력서\AVEVA - Marine Principal Technical Support & Consultant – PLM SME, Busan\Proposal\Future_Industrial_PLM_Meeting_Deck_EN_Rebuilt_v2.pptx"
$name = "Future_Industrial_PLM_Meeting_Deck_EN_Rebuilt_v2.pptx"
$outDir = "C:\Users\namma\.claude\plm_slide_work\review_png5"
New-Item -ItemType Directory -Force -Path $outDir | Out-Null

function Rgb([string]$hex){ $hex=$hex.TrimStart('#'); [int]([Convert]::ToInt32($hex.Substring(0,2),16)+[Convert]::ToInt32($hex.Substring(2,2),16)*256+[Convert]::ToInt32($hex.Substring(4,2),16)*65536) }
function Retry($sb){ for($a=0;$a -lt 6;$a++){ try{ return (& $sb) }catch{ Start-Sleep -Milliseconds 80 } }; return (& $sb) }

$WHITE="FFFFFF"; $LIGHT="C7D6E5"; $MUTE="93A8C0"
$CYAN="34C0EE"; $BLUE="4F8BE8"; $GREEN="49C28C"; $ORANGE="F0832F"; $VIOLET="9B7BE0"; $MAG="D957B8"
$BGD="0B1626"; $FRAME="0C1A2C"; $FRLINE="2C4663"; $BARC="15273C"; $HEXBG="241038"
$SHEET="0C2233"; $SHLINE="2E6E8E"; $VIEW="3FA9D6"; $CODEBG="0A1626"; $CODELN="27425C"; $SLATE="1A2740"

$ppt = New-Object -ComObject PowerPoint.Application
foreach($p in @($ppt.Presentations)){ if($p.Name -eq $name){ try{ $p.Saved=$true; $p.Close() }catch{} } }
$pres=$ppt.Presentations.Open($v2,0,0,0)

function NoShadow($s){ try{$s.Shadow.Visible=0}catch{} }
function Panel($l,$t,$w,$h,$fill,$line,$lw,$trans,$rad){ $s=$script:slide.Shapes.AddShape(5,[single]$l,[single]$t,[single]$w,[single]$h); try{$s.Adjustments.Item(1)=[single]$rad}catch{}; $s.Fill.Solid(); Retry{$s.Fill.ForeColor.RGB=[int](Rgb $fill)}; $s.Fill.Transparency=[single]$trans; if($line){$s.Line.Visible=-1; Retry{$s.Line.ForeColor.RGB=[int](Rgb $line)}; $s.Line.Weight=[single]$lw}else{$s.Line.Visible=0}; NoShadow $s; return $s }
function Bar($l,$t,$w,$h,$fill,$rad){ $s=$script:slide.Shapes.AddShape(5,[single]$l,[single]$t,[single]$w,[single]$h); try{$s.Adjustments.Item(1)=[single]$rad}catch{}; $s.Fill.Solid(); Retry{$s.Fill.ForeColor.RGB=[int](Rgb $fill)}; $s.Line.Visible=0; NoShadow $s; return $s }
function RectO($l,$t,$w,$h,$line,$lw){ $s=$script:slide.Shapes.AddShape(1,[single]$l,[single]$t,[single]$w,[single]$h); $s.Fill.Visible=0; $s.Line.Visible=-1; Retry{$s.Line.ForeColor.RGB=[int](Rgb $line)}; $s.Line.Weight=[single]$lw; NoShadow $s; return $s }
function Oval($l,$t,$w,$h,$fill){ $s=$script:slide.Shapes.AddShape(9,[single]$l,[single]$t,[single]$w,[single]$h); $s.Fill.Solid(); Retry{$s.Fill.ForeColor.RGB=[int](Rgb $fill)}; $s.Line.Visible=0; NoShadow $s; return $s }
function Txt($l,$t,$w,$h,$text,$size,$bold,$color,$font,$align,$anchor){ $tb=$script:slide.Shapes.AddTextbox(1,[single]$l,[single]$t,[single]$w,[single]$h); $tb.Fill.Visible=0;$tb.Line.Visible=0;NoShadow $tb; $tf=$tb.TextFrame;$tf.WordWrap=-1;$tf.AutoSize=0;$tf.MarginLeft=[single]2;$tf.MarginRight=[single]2;$tf.MarginTop=[single]1;$tf.MarginBottom=[single]1;$tf.VerticalAnchor=[int]$anchor; $tr=$tf.TextRange;$tr.Text=$text;$tr.Font.Size=[single]$size;$tr.Font.Bold=[int]$bold;$tr.Font.Name=[string]$font;Retry{$tr.Font.Color.RGB=[int](Rgb $color)};$tr.ParagraphFormat.Alignment=[int]$align; try{$tr.ParagraphFormat.SpaceBefore=[single]0;$tr.ParagraphFormat.SpaceAfter=[single]0}catch{}; return $tb }
function Chip($l,$t,$w,$h,$text,$size,$fill,$tcolor,$line,$font){ $s=$script:slide.Shapes.AddShape(5,[single]$l,[single]$t,[single]$w,[single]$h); try{$s.Adjustments.Item(1)=[single]0.4}catch{}; $s.Fill.Solid();Retry{$s.Fill.ForeColor.RGB=[int](Rgb $fill)}; if($line){$s.Line.Visible=-1;Retry{$s.Line.ForeColor.RGB=[int](Rgb $line)};$s.Line.Weight=[single]0.75}else{$s.Line.Visible=0}; NoShadow $s; $tf=$s.TextFrame;$tf.WordWrap=-1;$tf.AutoSize=0;$tf.VerticalAnchor=[int]3;$tf.MarginLeft=[single]3;$tf.MarginRight=[single]3;$tf.MarginTop=[single]0;$tf.MarginBottom=[single]0; $tr=$tf.TextRange;$tr.Text=$text;$tr.Font.Size=[single]$size;$tr.Font.Name=[string]$font;Retry{$tr.Font.Color.RGB=[int](Rgb $tcolor)};$tr.ParagraphFormat.Alignment=[int]2; return $s }
function Arrow($x1,$y1,$x2,$y2,$color,$w){ $c=$script:slide.Shapes.AddConnector(1,[single]$x1,[single]$y1,[single]$x2,[single]$y2); Retry{$c.Line.ForeColor.RGB=[int](Rgb $color)};$c.Line.Weight=[single]$w;$c.Line.BeginArrowheadStyle=[int]1;$c.Line.EndArrowheadStyle=[int]2;NoShadow $c; return $c }
function PLine($x1,$y1,$x2,$y2,$color,$w){ $c=$script:slide.Shapes.AddConnector(1,[single]$x1,[single]$y1,[single]$x2,[single]$y2); Retry{$c.Line.ForeColor.RGB=[int](Rgb $color)};$c.Line.Weight=[single]$w;$c.Line.BeginArrowheadStyle=[int]1;$c.Line.EndArrowheadStyle=[int]1;NoShadow $c; return $c }
function Hex($cx,$cy,$w,$h,$fill,$line,$text,$tcolor,$size){ $s=$script:slide.Shapes.AddShape(10,[single]($cx-$w/2),[single]($cy-$h/2),[single]$w,[single]$h); $s.Fill.Solid();Retry{$s.Fill.ForeColor.RGB=[int](Rgb $fill)};$s.Line.Visible=-1;Retry{$s.Line.ForeColor.RGB=[int](Rgb $line)};$s.Line.Weight=[single]1.25;NoShadow $s; $tf=$s.TextFrame;$tf.WordWrap=-1;$tf.AutoSize=0;$tf.VerticalAnchor=[int]3;$tf.MarginLeft=[single]1;$tf.MarginRight=[single]1; $tr=$tf.TextRange;$tr.Text=$text;$tr.Font.Size=[single]$size;$tr.Font.Bold=[int](-1);$tr.Font.Name="Aptos";Retry{$tr.Font.Color.RGB=[int](Rgb $tcolor)};$tr.ParagraphFormat.Alignment=[int]2; return $s }
function FindLike($pres,$prefix){ foreach($sl in $pres.Slides){ foreach($sh in $sl.Shapes){ try{ if($sh.HasTextFrame -eq -1 -and $sh.TextFrame.HasText -eq -1){ if($sh.TextFrame.TextRange.Text.Trim().StartsWith($prefix)){ return $sl } } }catch{} } }; return $null }

$flow = FindLike $pres "AI-gated"
$ins = [int]($flow.SlideIndex + 1)
$script:slide = $pres.Slides.Add($ins, 12)
Write-Output ("NEW SCREEN SLIDE idx=" + $ins)

# background + header
Bar 0 0 960 540 $BGD 0 | Out-Null
Txt 26 10 760 26 "Design tool -> 2D production drawing (AutoCAD DWG/DXF) -> registered in the API-built PLM" 16 -1 $WHITE "Aptos Display" 1 1 | Out-Null
Txt 770 12 158 14 "Future Industrial PLM" 8.5 0 $MUTE "Aptos" 3 1 | Out-Null
Txt 905 28 30 14 ([string]$ins) 9 0 $MUTE "Aptos" 3 1 | Out-Null
Txt 26 38 908 16 "AVEVA E3D/Marine model -> AVEVA Draw issues the 2D sheet -> export AutoCAD DWG/DXF (+MTO) -> AI checks drawing standards -> registered & status-tracked in the PLM." 9 0 $LIGHT "Aptos" 1 1 | Out-Null

# ===== LEFT: AVEVA Draw application window =====
Panel 22 64 590 410 $FRAME $FRLINE 1 0 0.02 | Out-Null
Bar 22 64 590 22 $BARC 0.04 | Out-Null
Oval 34 71 8 8 "FF5F56" | Out-Null
Oval 46 71 8 8 "FFBD2E" | Out-Null
Oval 58 71 8 8 "27C93F" | Out-Null
Txt 74 64 530 22 "AVEVA E3D Design  —  Draw (Draughting)   ·   Hull / Block B12 / Panel P3" 8 0 $MUTE "Consolas" 1 3 | Out-Null

# model tree panel
Panel 30 92 144 374 "0E1F33" "23456A" 0.75 0 0.04 | Out-Null
Txt 38 96 130 14 "MODEL TREE  (E3D)" 7.5 -1 $MUTE "Aptos" 1 1 | Out-Null
Txt 40 116 130 13 "Hull" 8 0 $LIGHT "Aptos" 1 1 | Out-Null
Txt 50 132 124 13 "Block B12" 8 -1 $CYAN "Aptos" 1 1 | Out-Null
Txt 60 148 120 13 "Panel P3" 8 -1 $CYAN "Aptos" 1 1 | Out-Null
Txt 60 164 120 13 "Stiffeners" 8 0 $LIGHT "Aptos" 1 1 | Out-Null
Txt 60 180 120 13 "Brackets" 8 0 $LIGHT "Aptos" 1 1 | Out-Null
Txt 50 196 124 13 "Block B13" 8 0 $LIGHT "Aptos" 1 1 | Out-Null
# tiny 3D wireframe mock
RectO 52 250 64 40 $VIEW 1 | Out-Null
RectO 66 240 64 40 $VIEW 1 | Out-Null
PLine 52 250 66 240 $VIEW 1 | Out-Null
PLine 116 250 130 240 $VIEW 1 | Out-Null
PLine 52 290 66 280 $VIEW 1 | Out-Null
PLine 116 290 130 280 $VIEW 1 | Out-Null
Txt 38 296 130 13 "3D model (E3D / Marine)" 7 0 $MUTE "Aptos" 2 1 | Out-Null

# drawing sheet
Panel 184 96 414 330 $SHEET $SHLINE 1 0 0.01 | Out-Null
# dimension line (top)
PLine 224 118 392 118 $VIEW 0.75 | Out-Null
PLine 224 114 224 122 $VIEW 0.75 | Out-Null
PLine 392 114 392 122 $VIEW 0.75 | Out-Null
Txt 224 106 168 12 "12 000" 7 0 $VIEW "Consolas" 2 1 | Out-Null
# view 1 (plan)
RectO 224 126 168 92 $VIEW 1 | Out-Null
PLine 224 150 392 150 $VIEW 0.5 | Out-Null
PLine 224 174 392 174 $VIEW 0.5 | Out-Null
PLine 224 198 392 198 $VIEW 0.5 | Out-Null
Txt 226 128 168 12 "PLAN" 6.5 0 $MUTE "Consolas" 1 1 | Out-Null
# view 2 (elevation)
RectO 224 232 168 70 $VIEW 1 | Out-Null
PLine 224 268 392 268 $VIEW 0.5 | Out-Null
Txt 226 234 168 12 "ELEVATION" 6.5 0 $MUTE "Consolas" 1 1 | Out-Null
# view 3 (iso/detail)
RectO 410 126 170 120 $VIEW 1 | Out-Null
PLine 410 126 580 246 $VIEW 0.5 | Out-Null
PLine 580 126 410 246 $VIEW 0.5 | Out-Null
Txt 412 128 168 12 "ISO / DETAIL" 6.5 0 $MUTE "Consolas" 1 1 | Out-Null
# title block
Panel 410 312 168 104 "0A1A28" $SHLINE 0.75 0 0.02 | Out-Null
Txt 418 316 152 14 "AVEVA Draw" 8 -1 $VIEW "Aptos" 1 1 | Out-Null
Txt 418 334 152 12 "Block B12 · Panel P3" 7.5 0 $LIGHT "Consolas" 1 1 | Out-Null
Txt 418 350 152 12 "Sheet 1/3 · 1:50 · A1" 7.5 0 $LIGHT "Consolas" 1 1 | Out-Null
Txt 418 366 152 12 "rev A · format DWG" 7.5 0 $LIGHT "Consolas" 1 1 | Out-Null
Txt 418 388 152 14 "drawing-standard: OK" 7.5 -1 $GREEN "Aptos" 1 1 | Out-Null
# export action bar
Chip 184 440 244 24 "Export  ->  AutoCAD DWG / DXF" 9 "1E6F8E" $WHITE $CYAN "Aptos" | Out-Null
Chip 436 440 162 24 "A1 · 1:50 · layers mapped" 7.5 "12273C" $MUTE "2C4663" "Aptos" | Out-Null

# ===== RIGHT: register in API-built PLM =====
Txt 628 90 312 16 "Registered in the API-built PLM" 11 -1 $ORANGE "Aptos" 1 1 | Out-Null
Chip 636 112 294 24 "B12-P3_revA.dwg     ·     + MTO" 8.5 "15273C" $CYAN $SHLINE "Consolas" | Out-Null
Arrow 783 138 783 150 $MUTE 1.5 | Out-Null
Hex 668 170 44 38 $HEXBG $MAG "AI" $MAG 9 | Out-Null
Txt 698 156 236 28 "drawing & standard check: title block, layers, refs" 8 0 $MAG "Aptos" 1 3 | Out-Null
Arrow 783 192 783 204 $MUTE 1.5 | Out-Null
# API window
Panel 636 206 294 96 $CODEBG $CODELN 1 0 0.04 | Out-Null
Txt 644 208 280 12 "POST  /api/v1/drawings" 7.5 -1 $CYAN "Consolas" 1 1 | Out-Null
$code = @('{ "block":"B12", "sheet":"P3",','  "rev":"A", "format":"DWG",','  "file":"B12-P3_revA.dwg" }','-> 201  drawing_id: 8842') -join "`r"
$cb = Txt 644 222 280 76 $code 7.5 0 "D7E7FF" "Consolas" 1 1
try{ $cb.TextFrame.TextRange.Paragraphs(4,1).Font.Color.RGB=[int](Rgb $GREEN) }catch{}
Arrow 783 304 783 316 $MUTE 1.5 | Out-Null
Chip 636 318 294 22 "registered   OK" 8.5 "15402F" $GREEN $GREEN "Aptos" | Out-Null
Chip 636 344 294 22 "linked to Block B12 / Panel P3" 8 $SLATE $LIGHT "2C4663" "Aptos" | Out-Null
Chip 636 370 294 22 "rev A · current · status-tracked" 8 $SLATE $LIGHT "2C4663" "Aptos" | Out-Null

# caption
Txt 26 482 908 32 "AVEVA Draw issues the 2D production drawing (DWG/DXF) from the E3D/Marine model; the API-built PLM registers it with its MTO, AI-checks drawing standards, and tracks revision & status across the digital thread." 8.5 0 $LIGHT "Aptos" 1 1 | Out-Null

$pres.Save()
$pres.Slides.Item($ins).Export("$outDir\screen.png","PNG",1280,720)
$pres.Close()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($pres) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($ppt)  | Out-Null
[GC]::Collect(); [GC]::WaitForPendingFinalizers()
Write-Output ("BUILD5B_DONE idx=" + $ins)
