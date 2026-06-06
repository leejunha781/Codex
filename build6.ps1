$ErrorActionPreference = "Stop"
$v2 = "D:\이력서\AVEVA - Marine Principal Technical Support & Consultant – PLM SME, Busan\Proposal\Future_Industrial_PLM_Meeting_Deck_EN_Rebuilt_v2.pptx"
$name = "Future_Industrial_PLM_Meeting_Deck_EN_Rebuilt_v2.pptx"
$outDir = "C:\Users\namma\.claude\plm_slide_work\review_png6"
New-Item -ItemType Directory -Force -Path $outDir | Out-Null

function Rgb([string]$hex){ $hex=$hex.TrimStart('#'); [int]([Convert]::ToInt32($hex.Substring(0,2),16)+[Convert]::ToInt32($hex.Substring(2,2),16)*256+[Convert]::ToInt32($hex.Substring(4,2),16)*65536) }
function Retry($sb){ for($a=0;$a -lt 6;$a++){ try{ return (& $sb) }catch{ Start-Sleep -Milliseconds 80 } }; return (& $sb) }

$WHITE="FFFFFF"; $LIGHT="C7D6E5"; $MUTE="93A8C0"
$CYAN="34C0EE"; $BLUE="4F8BE8"; $GREEN="49C28C"; $ORANGE="F0832F"; $VIOLET="9B7BE0"; $MAG="D957B8"; $AMBER="E8B23A"
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
function FindLike($pres,$prefix){ foreach($sl in $pres.Slides){ foreach($sh in $sl.Shapes){ try{ if($sh.HasTextFrame -eq -1 -and $sh.TextFrame.HasText -eq -1){ if($sh.TextFrame.TextRange.Text.Trim().StartsWith($prefix)){ return $sl } } }catch{} } }; return $null }

$screen = FindLike $pres "Design tool"
$ins = if($screen -ne $null){ [int]($screen.SlideIndex+1) } else { [int]($pres.Slides.Count+1) }
$script:slide = $pres.Slides.Add($ins, 12)
Write-Output ("NEW VCRM SLIDE idx=" + $ins)

# bg + header
Bar 0 0 960 540 $BGD 0 | Out-Null
Txt 26 10 730 26 "Requirement -> AI-built VCRM & Checklist, approved before production" 17 -1 $WHITE "Aptos Display" 1 1 | Out-Null
Txt 770 12 158 14 "Future Industrial PLM" 8.5 0 $MUTE "Aptos" 3 1 | Out-Null
Txt 905 28 30 14 ([string]$ins) 9 0 $MUTE "Aptos" 3 1 | Out-Null
Txt 26 38 908 16 "Enter a requirement in the PLM UI -> AI builds the VCRM + development Checklist -> attach deliverables per sub-item -> VCRM & Checklist are approval items -> approve, then release to production." 8.5 0 $LIGHT "Aptos" 1 1 | Out-Null

# window
Panel 22 62 916 454 $FRAME $FRLINE 1 0 0.02 | Out-Null
Bar 22 62 916 22 $BARC 0.04 | Out-Null
Oval 34 69 8 8 "FF5F56" | Out-Null
Oval 46 69 8 8 "FFBD2E" | Out-Null
Oval 58 69 8 8 "27C93F" | Out-Null
Txt 74 62 520 22 "Future PLM  —  Requirements & Verification   ·   /api/v1" 8 0 $MUTE "Consolas" 1 3 | Out-Null
Oval 866 69 8 8 $GREEN | Out-Null
Txt 878 62 56 22 "RUNNING" 8 -1 $GREEN "Aptos" 1 3 | Out-Null

# requirement input row
Panel 30 90 898 34 "10243A" $FRLINE 0.75 0 0.04 | Out-Null
Txt 40 90 84 34 "REQUIREMENT" 8.5 -1 $MUTE "Aptos" 1 3 | Out-Null
Panel 126 97 400 20 $INP $FRLINE 0.75 0 0.12 | Out-Null
Txt 132 97 392 20 "REQ-014 · Ballast Water Treatment · cap >= 250 m3/h · IMO D-2" 7.5 0 $LIGHT "Consolas" 1 3 | Out-Null
Chip 534 97 152 20 "POST /api/v1/requirements" 7 "12273C" $CYAN $FRLINE "Consolas" | Out-Null
Hex 708 107 36 30 $HEXBG $MAG "AI" $MAG 8.5 | Out-Null
Txt 730 90 196 34 "AI -> VCRM + Checklist + risk flags" 7.5 -1 $MAG "Aptos" 1 3 | Out-Null

# column headers
$LX=30; $LW=434; $RX=480; $RW=448
Bar $LX 130 $LW 26 $HDR 0.12 | Out-Null
Txt ($LX+10) 130 300 26 "VCRM — Verification Cross-Reference Matrix" 9.5 -1 $CYAN "Aptos" 1 3 | Out-Null
Txt ($LX+300) 130 124 26 "AI-gen · attach deliverables" 6.5 0 $MUTE "Aptos" 3 3 | Out-Null
Bar $RX 130 $RW 26 $HDR 0.12 | Out-Null
Txt ($RX+10) 130 320 26 "Checklist — Cautions & Requirement Reflection" 9.5 -1 $ORANGE "Aptos" 1 3 | Out-Null
Txt ($RX+320) 130 118 26 "AI-gen · attach deliverables" 6.5 0 $MUTE "Aptos" 3 3 | Out-Null

# VCRM sub-rows
$vcrm = @(
 @{id="1.1"; nm="Capacity >= 250 m3/h"; vf="verify: analysis/test"; dl="capacity_calc.pdf"; st="Verified"; sf="15402F"; sc=$GREEN},
 @{id="1.2"; nm="IMO D-2 compliance";   vf="verify: certification"; dl="IMO_D2_cert.pdf"; st="Certified"; sf="12233E"; sc=$BLUE},
 @{id="1.3"; nm="Power & energy budget"; vf="verify: test"; dl="power_test.pdf"; st="Verified"; sf="15402F"; sc=$GREEN},
 @{id="1.4"; nm="Piping / HVAC interface"; vf="verify: 3D / inspection"; dl="P&ID + 3D ref"; st="Linked"; sf="0E2A38"; sc=$CYAN},
 @{id="1.5"; nm="Maintenance access"; vf="verify: design review"; dl="access_review.docx"; st="Verified"; sf="15402F"; sc=$GREEN}
)
$rowY = @(160,213,266,319,372)
for($i=0;$i -lt 5;$i++){ $r=$vcrm[$i]; $y=$rowY[$i]
  Panel $LX $y $LW 48 $PANEL $ROWLN 0.75 0 0.07 | Out-Null
  Txt ($LX+12) ($y+4) 250 16 ($r.id+"   "+$r.nm) 9 -1 $WHITE "Aptos" 1 1 | Out-Null
  Txt ($LX+250) ($y+5) 172 13 $r.vf 7 0 $MUTE "Aptos" 3 1 | Out-Null
  Chip ($LX+12) ($y+25) 214 16 $r.dl 7.5 $SLATE "7FD3EE" $FRLINE "Consolas" | Out-Null
  Chip ($LX+330) ($y+25) 96 16 $r.st 8 $r.sf $r.sc $r.sc "Aptos" | Out-Null
}

# Checklist sub-rows
$chk = @(
 @{tag="Caution"; tf="3A2E10"; tc=$AMBER; nm="Class & IMO rules applied"; dl="rule_check.pdf"},
 @{tag="Caution"; tf="3A2E10"; tc=$AMBER; nm="Material / coating spec correct"; dl="material_spec.pdf"},
 @{tag="Caution"; tf="3A2E10"; tc=$AMBER; nm="3D clash check (interfaces)"; dl="clash_report.html"},
 @{tag="Reflect"; tf="0E2A38"; tc=$CYAN;  nm="All REQ items mapped to VCRM"; dl="trace_matrix.xlsx"},
 @{tag="Reflect"; tf="0E2A38"; tc=$CYAN;  nm="Test / inspection evidence"; dl="evidence_pack.zip"}
)
for($i=0;$i -lt 5;$i++){ $r=$chk[$i]; $y=$rowY[$i]
  Panel $RX $y $RW 48 $PANEL $ROWLN 0.75 0 0.07 | Out-Null
  Chip ($RX+10) ($y+5) 54 15 $r.tag 7 $r.tf $r.tc $r.tc "Aptos" | Out-Null
  Txt ($RX+72) ($y+4) 270 16 $r.nm 9 -1 $WHITE "Aptos" 1 1 | Out-Null
  Chip ($RX+12) ($y+25) 240 16 $r.dl 7.5 $SLATE "7FD3EE" $FRLINE "Consolas" | Out-Null
  Chip ($RX+352) ($y+25) 84 16 "OK" 8 "15402F" $GREEN $GREEN "Aptos" | Out-Null
}

# approval gate
Panel 30 460 898 52 "12243A" $ORANGE 1.25 0 0.04 | Out-Null
Txt 40 460 112 52 "FINAL APPROVAL  (production gate)" 9 -1 $WHITE "Aptos" 1 3 | Out-Null
Chip 160 468 152 22 "VCRM  5/5  OK" 8 "15402F" $GREEN $GREEN "Aptos" | Out-Null
Txt 314 468 12 22 "+" 9 -1 $MUTE "Aptos" 2 3 | Out-Null
Chip 330 468 140 22 "Checklist  OK" 8 "15402F" $GREEN $GREEN "Aptos" | Out-Null
Txt 472 468 12 22 "+" 9 -1 $MUTE "Aptos" 2 3 | Out-Null
Chip 488 468 200 22 "All requirements reflected  OK" 8 "15402F" $GREEN $GREEN "Aptos" | Out-Null
Arrow 694 479 716 479 $MUTE 1.75 | Out-Null
Chip 720 466 200 28 "PRODUCTION  RELEASED" 9.5 "1E7A4E" $WHITE $GREEN "Aptos" | Out-Null
Txt 160 494 760 14 "AI holds this gate until every VCRM line is verified and every checklist item is cleared — production cannot start otherwise." 7 0 $MUTE "Aptos" 1 1 | Out-Null

$pres.Save()
$pres.Slides.Item($ins).Export("$outDir\vcrm.png","PNG",1280,720)
$pres.Close()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($pres) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($ppt)  | Out-Null
[GC]::Collect(); [GC]::WaitForPendingFinalizers()
Write-Output ("BUILD6_DONE idx=" + $ins)
