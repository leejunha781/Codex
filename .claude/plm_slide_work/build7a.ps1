$ErrorActionPreference = "Stop"
$v2 = "D:\이력서\AVEVA - Marine Principal Technical Support & Consultant – PLM SME, Busan\Proposal\Future_Industrial_PLM_Meeting_Deck_EN_Rebuilt_v2.pptx"
$name = "Future_Industrial_PLM_Meeting_Deck_EN_Rebuilt_v2.pptx"
$outDir = "C:\Users\namma\.claude\plm_slide_work\review_png7"
New-Item -ItemType Directory -Force -Path $outDir | Out-Null

function Rgb([string]$hex){ $hex=$hex.TrimStart('#'); [int]([Convert]::ToInt32($hex.Substring(0,2),16)+[Convert]::ToInt32($hex.Substring(2,2),16)*256+[Convert]::ToInt32($hex.Substring(4,2),16)*65536) }
function Retry($sb){ for($a=0;$a -lt 6;$a++){ try{ return (& $sb) }catch{ Start-Sleep -Milliseconds 80 } }; return (& $sb) }

$WHITE="FFFFFF"; $LIGHT="C7D6E5"; $MUTE="93A8C0"
$CYAN="34C0EE"; $BLUE="4F8BE8"; $GREEN="49C28C"; $ORANGE="F0832F"; $VIOLET="9B7BE0"; $MAG="D957B8"; $AMBER="E8B23A"; $RED="E5634E"
$BGD="0B1626"; $FRAME="0C1A2C"; $FRLINE="2C4663"; $BARC="15273C"; $HEXBG="241038"
$PANEL="14253C"; $ROWLN="243B57"; $HDR="15314A"; $SLATE="1A2740"; $INP="0E1F33"

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
function Hex($cx,$cy,$w,$h,$fill,$line,$text,$tcolor,$size){ $s=$script:slide.Shapes.AddShape(10,[single]($cx-$w/2),[single]($cy-$h/2),[single]$w,[single]$h); $s.Fill.Solid();Retry{$s.Fill.ForeColor.RGB=[int](Rgb $fill)};$s.Line.Visible=-1;Retry{$s.Line.ForeColor.RGB=[int](Rgb $line)};$s.Line.Weight=[single]1.25;NoShadow $s; $tf=$s.TextFrame;$tf.WordWrap=-1;$tf.AutoSize=0;$tf.VerticalAnchor=[int]3; $tr=$tf.TextRange;$tr.Text=$text;$tr.Font.Size=[single]$size;$tr.Font.Bold=[int](-1);$tr.Font.Name="Aptos";Retry{$tr.Font.Color.RGB=[int](Rgb $tcolor)};$tr.ParagraphFormat.Alignment=[int]2; return $s }
function DelAll($sl){ foreach($sh in @($sl.Shapes)){ try{$sh.Delete()}catch{} } }
function FindLike($pres,$prefix){ foreach($sl in $pres.Slides){ foreach($sh in $sl.Shapes){ try{ if($sh.HasTextFrame -eq -1 -and $sh.TextFrame.HasText -eq -1){ if($sh.TextFrame.TextRange.Text.Trim().StartsWith($prefix)){ return $sl } } }catch{} } }; return $null }
function SetPageNum($sl,$num){ foreach($sh in $sl.Shapes){ try{ if($sh.HasTextFrame -eq -1 -and $sh.TextFrame.HasText -eq -1){ $t=$sh.TextFrame.TextRange.Text.Trim(); if($sh.Left -ge 895 -and $t -match '^[0-9]{1,3}$'){ $sh.TextFrame.TextRange.Text=[string]$num; return } } }catch{} } }
function EditText($sl,$old,$new){ foreach($sh in $sl.Shapes){ try{ if($sh.HasTextFrame -eq -1 -and $sh.TextFrame.HasText -eq -1){ if($sh.TextFrame.TextRange.Text.Trim() -eq $old){ $sh.TextFrame.TextRange.Text=$new; $r=$sh.TextFrame.TextRange; $r.Font.Size=[single]8; $r.Font.Name="Aptos"; Retry{$r.Font.Color.RGB=[int](Rgb $LIGHT)}; return $true } } }catch{} }; return $false }

# ============ regenerate VCRM slide ============
$vcrm = FindLike $pres "Requirement ->"
Write-Output ("VCRM found at idx=" + $vcrm.SlideIndex)
DelAll $vcrm
$script:slide = $vcrm

Bar 0 0 960 540 $BGD 0 | Out-Null
Txt 26 10 660 26 "Requirement -> AI-built VCRM & Checklist (admin-editable, approved in Impact+Approval)" 15 -1 $WHITE "Aptos Display" 1 1 | Out-Null
Txt 770 12 158 14 "Future Industrial PLM" 8.5 0 $MUTE "Aptos" 3 1 | Out-Null
Txt 905 28 30 14 "12" 9 0 $MUTE "Aptos" 3 1 | Out-Null
Txt 26 38 908 16 "Enter a requirement -> AI builds the VCRM + Checklist -> admin can add/delete sub-items -> attach deliverables -> VCRM & Checklist are approved together in the Impact + Approval step before production." 8.5 0 $LIGHT "Aptos" 1 1 | Out-Null

Panel 22 62 916 454 $FRAME $FRLINE 1 0 0.02 | Out-Null
Bar 22 62 916 22 $BARC 0.04 | Out-Null
Oval 34 69 8 8 "FF5F56" | Out-Null
Oval 46 69 8 8 "FFBD2E" | Out-Null
Oval 58 69 8 8 "27C93F" | Out-Null
Txt 74 62 460 22 "Future PLM  —  Requirements & Verification   ·   /api/v1" 8 0 $MUTE "Consolas" 1 3 | Out-Null
Chip 612 67 70 14 "ADMIN" 7 "3A2E10" $AMBER $AMBER "Aptos" | Out-Null
Oval 866 69 8 8 $GREEN | Out-Null
Txt 878 62 56 22 "RUNNING" 8 -1 $GREEN "Aptos" 1 3 | Out-Null

# requirement input row
Panel 30 88 898 32 "10243A" $FRLINE 0.75 0 0.04 | Out-Null
Txt 40 88 84 32 "REQUIREMENT" 8.5 -1 $MUTE "Aptos" 1 3 | Out-Null
Panel 126 95 396 18 $INP $FRLINE 0.75 0 0.12 | Out-Null
Txt 132 95 388 18 "REQ-014 · Ballast Water Treatment · cap >= 250 m3/h · IMO D-2" 7.5 0 $LIGHT "Consolas" 1 3 | Out-Null
Chip 530 95 150 18 "POST /api/v1/requirements" 7 "12273C" $CYAN $FRLINE "Consolas" | Out-Null
Hex 706 104 34 28 $HEXBG $MAG "AI" $MAG 8.5 | Out-Null
Txt 728 88 200 32 "AI -> VCRM + Checklist + risk flags" 7.5 -1 $MAG "Aptos" 1 3 | Out-Null

$LX=30; $LW=434; $RX=480; $RW=448
# headers
Bar $LX 126 $LW 24 $HDR 0.12 | Out-Null
Txt ($LX+10) 126 380 24 "VCRM — Verification Cross-Reference Matrix  (대분류)" 9 -1 $CYAN "Aptos" 1 3 | Out-Null
Bar $RX 126 $RW 24 $HDR 0.12 | Out-Null
Txt ($RX+10) 126 400 24 "Checklist — Cautions & Requirement Reflection  (대분류)" 9 -1 $ORANGE "Aptos" 1 3 | Out-Null
# admin toolbars
Chip $LX 154 70 17 "+ Add item" 7 "10303A" $CYAN $CYAN "Aptos" | Out-Null
Chip ($LX+76) 154 64 17 "× Delete" 7 "301412" $RED $RED "Aptos" | Out-Null
Txt ($LX+146) 153 288 18 "admin · POST / DELETE  /api/v1/vcrm/items" 6.5 0 $MUTE "Consolas" 3 3 | Out-Null
Chip $RX 154 70 17 "+ Add item" 7 "10303A" $CYAN $CYAN "Aptos" | Out-Null
Chip ($RX+76) 154 64 17 "× Delete" 7 "301412" $RED $RED "Aptos" | Out-Null
Txt ($RX+146) 153 302 18 "admin · POST / DELETE  /api/v1/checklist/items" 6.5 0 $MUTE "Consolas" 3 3 | Out-Null

# VCRM rows
$vc = @(
 @{id="1.1"; nm="Capacity >= 250 m3/h"; vf="verify: analysis/test"; dl="capacity_calc.pdf"; st="Verified"; sf="15402F"; sc=$GREEN},
 @{id="1.2"; nm="IMO D-2 compliance"; vf="verify: certification"; dl="IMO_D2_cert.pdf"; st="Certified"; sf="12233E"; sc=$BLUE},
 @{id="1.3"; nm="Power & energy budget"; vf="verify: test"; dl="power_test.pdf"; st="Verified"; sf="15402F"; sc=$GREEN},
 @{id="1.4"; nm="Piping / HVAC interface"; vf="verify: 3D/inspection"; dl="P&ID + 3D ref"; st="Linked"; sf="0E2A38"; sc=$CYAN},
 @{id="1.5"; nm="Maintenance access"; vf="verify: design review"; dl="access_review.docx"; st="Verified"; sf="15402F"; sc=$GREEN}
)
$rowY=@(178,226,274,322,370)
for($i=0;$i -lt 5;$i++){ $r=$vc[$i]; $y=$rowY[$i]
  Panel $LX $y $LW 44 $PANEL $ROWLN 0.75 0 0.08 | Out-Null
  Txt ($LX+12) ($y+3) 240 15 ($r.id+"   "+$r.nm) 8.5 -1 $WHITE "Aptos" 1 1 | Out-Null
  Txt ($LX+250) ($y+4) 150 12 $r.vf 6.5 0 $MUTE "Aptos" 3 1 | Out-Null
  Txt ($LX+416) ($y+2) 14 14 "×" 9 0 $MUTE "Aptos" 2 1 | Out-Null
  Chip ($LX+12) ($y+22) 210 16 $r.dl 7.5 $SLATE "7FD3EE" $FRLINE "Consolas" | Out-Null
  Chip ($LX+330) ($y+22) 96 16 $r.st 8 $r.sf $r.sc $r.sc "Aptos" | Out-Null
}
# Checklist rows
$ck = @(
 @{tag="Caution"; tf="3A2E10"; tc=$AMBER; nm="Class & IMO rules applied"; dl="rule_check.pdf"},
 @{tag="Caution"; tf="3A2E10"; tc=$AMBER; nm="Material / coating spec correct"; dl="material_spec.pdf"},
 @{tag="Caution"; tf="3A2E10"; tc=$AMBER; nm="3D clash check (interfaces)"; dl="clash_report.html"},
 @{tag="Reflect"; tf="0E2A38"; tc=$CYAN; nm="All REQ items mapped to VCRM"; dl="trace_matrix.xlsx"},
 @{tag="Reflect"; tf="0E2A38"; tc=$CYAN; nm="Test / inspection evidence"; dl="evidence_pack.zip"}
)
for($i=0;$i -lt 5;$i++){ $r=$ck[$i]; $y=$rowY[$i]
  Panel $RX $y $RW 44 $PANEL $ROWLN 0.75 0 0.08 | Out-Null
  Chip ($RX+10) ($y+4) 52 14 $r.tag 7 $r.tf $r.tc $r.tc "Aptos" | Out-Null
  Txt ($RX+70) ($y+3) 250 15 $r.nm 8.5 -1 $WHITE "Aptos" 1 1 | Out-Null
  Txt ($RX+430) ($y+2) 14 14 "×" 9 0 $MUTE "Aptos" 2 1 | Out-Null
  Chip ($RX+12) ($y+22) 230 16 $r.dl 7.5 $SLATE "7FD3EE" $FRLINE "Consolas" | Out-Null
  Chip ($RX+352) ($y+22) 84 16 "OK" 8 "15402F" $GREEN $GREEN "Aptos" | Out-Null
}

# approval gate (in Impact + Approval step)
Panel 30 458 898 52 "12243A" $VIOLET 1.25 0 0.04 | Out-Null
Txt 40 458 118 52 "APPROVED IN THE  'IMPACT + APPROVAL' STEP" 8.5 -1 $VIOLET "Aptos" 1 3 | Out-Null
Chip 168 466 150 22 "VCRM  5/5  OK" 8 "15402F" $GREEN $GREEN "Aptos" | Out-Null
Txt 320 466 12 22 "+" 9 -1 $MUTE "Aptos" 2 3 | Out-Null
Chip 336 466 138 22 "Checklist  OK" 8 "15402F" $GREEN $GREEN "Aptos" | Out-Null
Txt 476 466 12 22 "+" 9 -1 $MUTE "Aptos" 2 3 | Out-Null
Chip 492 466 196 22 "All requirements reflected  OK" 8 "15402F" $GREEN $GREEN "Aptos" | Out-Null
Arrow 694 477 716 477 $MUTE 1.75 | Out-Null
Chip 720 464 200 28 "PRODUCTION  RELEASED" 9.5 "1E7A4E" $WHITE $GREEN "Aptos" | Out-Null
Txt 168 492 760 14 "VCRM & Checklist are approval items inside the Impact + Approval step — production releases only when all are approved together." 7 0 $MUTE "Aptos" 1 1 | Out-Null

# ============ edit flow slide's Impact+Approval card ============
$flow = FindLike $pres "AI-gated"
$ok = EditText $flow "Affected objects + approval" "Affected objects + VCRM & Checklist approval"
Write-Output ("flow card edited=" + $ok + " at idx=" + $flow.SlideIndex)

# ============ move VCRM before flow (to position 12) ============
$vcrm.MoveTo(12)
Write-Output ("after move: VCRM idx=" + $vcrm.SlideIndex + "  flow idx=" + $flow.SlideIndex)

# fix page numbers on the 3 reordered slides
SetPageNum $vcrm 12
SetPageNum $flow 13
$draw = FindLike $pres "Design tool"
if($draw -ne $null){ SetPageNum $draw 14 }

$pres.Save()
$pres.Slides.Item($vcrm.SlideIndex).Export("$outDir\vcrm.png","PNG",1280,720)
$pres.Slides.Item($flow.SlideIndex).Export("$outDir\flow.png","PNG",1280,720)
$pres.Close()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($pres) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($ppt)  | Out-Null
[GC]::Collect(); [GC]::WaitForPendingFinalizers()
Write-Output "BUILD7A_DONE"
