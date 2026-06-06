$ErrorActionPreference = "Stop"
$v2 = "D:\이력서\AVEVA - Marine Principal Technical Support & Consultant – PLM SME, Busan\Proposal\Future_Industrial_PLM_Meeting_Deck_EN_Rebuilt_v2.pptx"
$name = "Future_Industrial_PLM_Meeting_Deck_EN_Rebuilt_v2.pptx"
$outDir = "C:\Users\namma\.claude\plm_slide_work\review_png7"

function Rgb([string]$hex){ $hex=$hex.TrimStart('#'); [int]([Convert]::ToInt32($hex.Substring(0,2),16)+[Convert]::ToInt32($hex.Substring(2,2),16)*256+[Convert]::ToInt32($hex.Substring(4,2),16)*65536) }
function Retry($sb){ for($a=0;$a -lt 6;$a++){ try{ return (& $sb) }catch{ Start-Sleep -Milliseconds 80 } }; return (& $sb) }
$WHITE="FFFFFF"; $LIGHT="C7D6E5"; $MUTE="93A8C0"
$CYAN="34C0EE"; $BLUE="4F8BE8"; $GREEN="49C28C"; $ORANGE="F0832F"; $VIOLET="9B7BE0"; $MAG="D957B8"; $AMBER="E8B23A"; $RED="E5634E"
$BGD="0B1626"; $FRAME="0C1A2C"; $FRLINE="2C4663"; $BARC="15273C"; $PANEL="14253C"; $ROWLN="243B57"; $CODEBG="0A1626"; $CODELN="27425C"; $SLATE="1A2740"

$ppt = New-Object -ComObject PowerPoint.Application
foreach($p in @($ppt.Presentations)){ if($p.Name -eq $name){ try{ $p.Saved=$true; $p.Close() }catch{} } }
$pres=$ppt.Presentations.Open($v2,0,0,0)

function NoShadow($s){ try{$s.Shadow.Visible=0}catch{} }
function Panel($l,$t,$w,$h,$fill,$line,$lw,$trans,$rad){ $s=$script:slide.Shapes.AddShape(5,[single]$l,[single]$t,[single]$w,[single]$h); try{$s.Adjustments.Item(1)=[single]$rad}catch{}; $s.Fill.Solid(); Retry{$s.Fill.ForeColor.RGB=[int](Rgb $fill)}; $s.Fill.Transparency=[single]$trans; if($line){$s.Line.Visible=-1; Retry{$s.Line.ForeColor.RGB=[int](Rgb $line)}; $s.Line.Weight=[single]$lw}else{$s.Line.Visible=0}; NoShadow $s; return $s }
function Bar($l,$t,$w,$h,$fill,$rad){ $s=$script:slide.Shapes.AddShape(5,[single]$l,[single]$t,[single]$w,[single]$h); try{$s.Adjustments.Item(1)=[single]$rad}catch{}; $s.Fill.Solid(); Retry{$s.Fill.ForeColor.RGB=[int](Rgb $fill)}; $s.Line.Visible=0; NoShadow $s; return $s }
function Oval($l,$t,$w,$h,$fill){ $s=$script:slide.Shapes.AddShape(9,[single]$l,[single]$t,[single]$w,[single]$h); $s.Fill.Solid(); Retry{$s.Fill.ForeColor.RGB=[int](Rgb $fill)}; $s.Line.Visible=0; NoShadow $s; return $s }
function Txt($l,$t,$w,$h,$text,$size,$bold,$color,$font,$align,$anchor){ $tb=$script:slide.Shapes.AddTextbox(1,[single]$l,[single]$t,[single]$w,[single]$h); $tb.Fill.Visible=0;$tb.Line.Visible=0;NoShadow $tb; $tf=$tb.TextFrame;$tf.WordWrap=-1;$tf.AutoSize=0;$tf.MarginLeft=[single]2;$tf.MarginRight=[single]2;$tf.MarginTop=[single]1;$tf.MarginBottom=[single]1;$tf.VerticalAnchor=[int]$anchor; $tr=$tf.TextRange;$tr.Text=$text;$tr.Font.Size=[single]$size;$tr.Font.Bold=[int]$bold;$tr.Font.Name=[string]$font;Retry{$tr.Font.Color.RGB=[int](Rgb $color)};$tr.ParagraphFormat.Alignment=[int]$align; try{$tr.ParagraphFormat.SpaceBefore=[single]0;$tr.ParagraphFormat.SpaceAfter=[single]0}catch{}; return $tb }
function Chip($l,$t,$w,$h,$text,$size,$fill,$tcolor,$line,$font){ $s=$script:slide.Shapes.AddShape(5,[single]$l,[single]$t,[single]$w,[single]$h); try{$s.Adjustments.Item(1)=[single]0.4}catch{}; $s.Fill.Solid();Retry{$s.Fill.ForeColor.RGB=[int](Rgb $fill)}; if($line){$s.Line.Visible=-1;Retry{$s.Line.ForeColor.RGB=[int](Rgb $line)};$s.Line.Weight=[single]0.75}else{$s.Line.Visible=0}; NoShadow $s; $tf=$s.TextFrame;$tf.WordWrap=-1;$tf.AutoSize=0;$tf.VerticalAnchor=[int]3;$tf.MarginLeft=[single]3;$tf.MarginRight=[single]3;$tf.MarginTop=[single]0;$tf.MarginBottom=[single]0; $tr=$tf.TextRange;$tr.Text=$text;$tr.Font.Size=[single]$size;$tr.Font.Name=[string]$font;Retry{$tr.Font.Color.RGB=[int](Rgb $tcolor)};$tr.ParagraphFormat.Alignment=[int]2; return $s }
function Arrow($x1,$y1,$x2,$y2,$color,$w){ $c=$script:slide.Shapes.AddConnector(1,[single]$x1,[single]$y1,[single]$x2,[single]$y2); Retry{$c.Line.ForeColor.RGB=[int](Rgb $color)};$c.Line.Weight=[single]$w;$c.Line.BeginArrowheadStyle=[int]1;$c.Line.EndArrowheadStyle=[int]2;NoShadow $c; return $c }
function FindLike($pres,$prefix){ foreach($sl in $pres.Slides){ foreach($sh in $sl.Shapes){ try{ if($sh.HasTextFrame -eq -1 -and $sh.TextFrame.HasText -eq -1){ if($sh.TextFrame.TextRange.Text.Trim().StartsWith($prefix)){ return $sl } } }catch{} } }; return $null }
function SetPageNum($sl,$num){ foreach($sh in $sl.Shapes){ try{ if($sh.HasTextFrame -eq -1 -and $sh.TextFrame.HasText -eq -1){ $t=$sh.TextFrame.TextRange.Text.Trim(); if($sh.Left -ge 895 -and $t -match '^[0-9]{1,3}$'){ $sh.TextFrame.TextRange.Text=[string]$num; return } } }catch{} } }

$draw = FindLike $pres "Design tool"
$ins = if($draw -ne $null){ [int]($draw.SlideIndex+1) } else { [int]($pres.Slides.Count+1) }
$script:slide = $pres.Slides.Add($ins, 12)
Write-Output ("NEW PROCESS SLIDE idx=" + $ins)

Bar 0 0 960 540 $BGD 0 | Out-Null
Txt 26 10 730 26 "User-configurable PLM process — reshape the whole flow via API" 17 -1 $WHITE "Aptos Display" 1 1 | Out-Null
Txt 770 12 158 14 "Future Industrial PLM" 8.5 0 $MUTE "Aptos" 3 1 | Out-Null
Txt 905 28 30 14 ([string]$ins) 9 0 $MUTE "Aptos" 3 1 | Out-Null
Txt 26 38 908 16 "Admins reorder, add or remove steps and reconfigure AI gates through governed API calls; the live PLM flow updates immediately — versioned, role-gated and audited." 8.5 0 $LIGHT "Aptos" 1 1 | Out-Null

# ===== LEFT: Process Designer window =====
Panel 22 62 452 454 $FRAME $FRLINE 1 0 0.02 | Out-Null
Bar 22 62 452 22 $BARC 0.04 | Out-Null
Oval 34 69 8 8 "FF5F56" | Out-Null
Oval 46 69 8 8 "FFBD2E" | Out-Null
Oval 58 69 8 8 "27C93F" | Out-Null
Txt 74 62 400 22 "Future PLM  —  Process Designer  ·  admin  ·  /api/v1" 8 0 $MUTE "Consolas" 1 3 | Out-Null
Txt 32 88 440 14 "Drag to reorder   ·   gate = configure AI gate   ·   x = remove   ·   + add step" 7 0 $MUTE "Aptos" 1 1 | Out-Null

$steps = @("Requirement","Design (E3D / Marine)","2D Drawing (AVEVA Draw)","E-BOM / MTO","Approval","Effectivity","Baseline","ECR / ECO","M-BOM / MCO","Impact + Approval")
$sy = 106
for($i=0;$i -lt $steps.Count;$i++){
  $isFirst = ($i -eq 0)
  $bc = if($isFirst){ $CYAN } else { $ROWLN }
  $dotc = if($i -lt 4){$CYAN}elseif($i -lt 7){$GREEN}elseif($i -lt 9){$ORANGE}else{$VIOLET}
  Panel 32 $sy 408 27 $PANEL $bc 0.75 0 0.12 | Out-Null
  Txt 40 ($sy+1) 16 25 "≡" 10 0 $MUTE "Aptos" 2 3 | Out-Null
  Oval 60 ($sy+8) 11 11 $dotc | Out-Null
  Txt 78 ($sy+1) 230 25 $steps[$i] 8.5 -1 $WHITE "Aptos" 1 3 | Out-Null
  Txt 300 ($sy+1) 134 25 "↑  ↓     gate     ×" 8 0 $MUTE "Aptos" 3 3 | Out-Null
  if($isFirst){ Chip 250 ($sy+5) 46 16 "moved" 7 "10303A" $CYAN $CYAN "Aptos" | Out-Null }
  $sy += 30
}
Chip 32 ($sy+2) 120 22 "+ Add step" 8 "10303A" $CYAN $CYAN "Aptos" | Out-Null
Chip 160 ($sy+2) 150 22 "Save & version process" 8 "15402F" $GREEN $GREEN "Aptos" | Out-Null

# ===== RIGHT TOP: API =====
Txt 486 86 452 16 "Process configuration API   (admin · role-gated)" 10.5 -1 $CYAN "Aptos" 1 1 | Out-Null
Panel 486 106 452 150 $CODEBG $CODELN 1 0 0.03 | Out-Null
Txt 496 110 432 14 "REST · OpenAPI 3.1 · OIDC" 7.5 -1 $MUTE "Consolas" 1 1 | Out-Null
$api = @(
 'GET     /api/v1/process/definition',
 'PUT     /api/v1/process/steps/{id}/order     # reorder',
 'POST    /api/v1/process/steps                # add step',
 'DELETE  /api/v1/process/steps/{id}           # remove',
 'PUT     /api/v1/process/gates/{id}           # AI gate'
) -join "`r"
Txt 496 128 432 120 $api 8.5 0 "D7E7FF" "Consolas" 1 1 | Out-Null
Txt 486 260 452 14 "versioned · role-gated (admin) · audit-logged · takes effect immediately" 7.5 0 $MUTE "Aptos" 1 1 | Out-Null

# ===== RIGHT BOTTOM: example change =====
Txt 486 282 452 16 "Example change: move Requirement before Design" 10.5 -1 $ORANGE "Aptos" 1 1 | Out-Null
Panel 486 302 452 96 $PANEL $ROWLN 0.75 0 0.05 | Out-Null
Txt 496 308 60 16 "BEFORE" 8 -1 $MUTE "Aptos" 1 3 | Out-Null
Chip 560 309 96 16 "Design (E3D)" 7.5 "0E2A38" $CYAN $FRLINE "Aptos" | Out-Null
Arrow 658 317 686 317 $MUTE 1.5 | Out-Null
Chip 690 309 100 16 "Requirement" 7.5 $SLATE $MUTE $FRLINE "Aptos" | Out-Null
Txt 794 308 130 16 "(req after design)" 7 0 $MUTE "Aptos" 1 3 | Out-Null
Chip 496 334 430 16 "PUT /api/v1/process/steps/requirement/order  { ""before"": ""design"" }" 7.5 "12273C" $MAG $CODELN "Consolas" | Out-Null
Txt 496 360 60 16 "AFTER" 8 -1 $GREEN "Aptos" 1 3 | Out-Null
Chip 560 361 100 16 "Requirement" 7.5 "0E2A38" $GREEN $GREEN "Aptos" | Out-Null
Arrow 662 369 690 369 $MUTE 1.5 | Out-Null
Chip 694 361 96 16 "Design (E3D)" 7.5 "0E2A38" $CYAN $FRLINE "Aptos" | Out-Null
Txt 794 360 130 16 "-> rest of flow" 7 0 $MUTE "Aptos" 1 3 | Out-Null

Txt 486 402 452 14 "Other changes via the same API:" 7.5 -1 $MUTE "Aptos" 1 1 | Out-Null
Chip 486 420 168 20 "POST /steps  ->  + Class Approval" 7 $SLATE $LIGHT $FRLINE "Consolas" | Out-Null
Chip 662 420 120 20 "DELETE /steps/{id}" 7 $SLATE $LIGHT $FRLINE "Consolas" | Out-Null
Chip 790 420 148 20 "PUT /gates  ->  threshold" 7 $SLATE $LIGHT $FRLINE "Consolas" | Out-Null
Txt 486 446 452 28 "AI re-validates the reconfigured process; every user immediately sees the updated flow, gates and live status — no redeploy, fully audited." 8 0 $LIGHT "Aptos" 1 1 | Out-Null

SetPageNum $script:slide $ins
$pres.Save()
$pres.Slides.Item($ins).Export("$outDir\process.png","PNG",1280,720)
$pres.Close()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($pres) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($ppt)  | Out-Null
[GC]::Collect(); [GC]::WaitForPendingFinalizers()
Write-Output ("BUILD7B_DONE idx=" + $ins)
