$ErrorActionPreference = "Stop"
$v2 = "D:\이력서\AVEVA - Marine Principal Technical Support & Consultant – PLM SME, Busan\Proposal\Future_Industrial_PLM_Meeting_Deck_EN_Rebuilt_v2.pptx"
$name = "Future_Industrial_PLM_Meeting_Deck_EN_Rebuilt_v2.pptx"
$outDir = "C:\Users\namma\.claude\plm_slide_work\review_png8"
function Rgb([string]$hex){ $hex=$hex.TrimStart('#'); [int]([Convert]::ToInt32($hex.Substring(0,2),16)+[Convert]::ToInt32($hex.Substring(2,2),16)*256+[Convert]::ToInt32($hex.Substring(4,2),16)*65536) }
function Retry($sb){ for($a=0;$a -lt 6;$a++){ try{ return (& $sb) }catch{ Start-Sleep -Milliseconds 80 } }; return (& $sb) }
$WHITE="FFFFFF"; $LIGHT="C7D6E5"; $MUTE="93A8C0"; $CYAN="34C0EE"; $GREEN="49C28C"; $ORANGE="F0832F"; $VIOLET="9B7BE0"; $SLATE2="8FA9C4"
$BGD="0B1626"; $PANEL="14253C"; $GRID="2C4663"
$ppt = New-Object -ComObject PowerPoint.Application
foreach($p in @($ppt.Presentations)){ if($p.Name -eq $name){ try{ $p.Saved=$true; $p.Close() }catch{} } }
$pres=$ppt.Presentations.Open($v2,0,0,0)
function NoShadow($s){ try{$s.Shadow.Visible=0}catch{} }
function Panel($l,$t,$w,$h,$fill,$line,$lw,$trans,$rad){ $s=$script:slide.Shapes.AddShape(5,[single]$l,[single]$t,[single]$w,[single]$h); try{$s.Adjustments.Item(1)=[single]$rad}catch{}; $s.Fill.Solid(); Retry{$s.Fill.ForeColor.RGB=[int](Rgb $fill)}; $s.Fill.Transparency=[single]$trans; if($line){$s.Line.Visible=-1; Retry{$s.Line.ForeColor.RGB=[int](Rgb $line)}; $s.Line.Weight=[single]$lw}else{$s.Line.Visible=0}; NoShadow $s; return $s }
function Bar($l,$t,$w,$h,$fill,$rad){ $s=$script:slide.Shapes.AddShape(5,[single]$l,[single]$t,[single]$w,[single]$h); try{$s.Adjustments.Item(1)=[single]$rad}catch{}; $s.Fill.Solid(); Retry{$s.Fill.ForeColor.RGB=[int](Rgb $fill)}; $s.Line.Visible=0; NoShadow $s; return $s }
function Oval($l,$t,$w,$h,$fill){ $s=$script:slide.Shapes.AddShape(9,[single]$l,[single]$t,[single]$w,[single]$h); $s.Fill.Solid(); Retry{$s.Fill.ForeColor.RGB=[int](Rgb $fill)}; $s.Line.Visible=0; NoShadow $s; return $s }
function Txt($l,$t,$w,$h,$text,$size,$bold,$color,$font,$align,$anchor){ $tb=$script:slide.Shapes.AddTextbox(1,[single]$l,[single]$t,[single]$w,[single]$h); $tb.Fill.Visible=0;$tb.Line.Visible=0;NoShadow $tb; $tf=$tb.TextFrame;$tf.WordWrap=-1;$tf.AutoSize=0;$tf.MarginLeft=[single]2;$tf.MarginRight=[single]2;$tf.MarginTop=[single]1;$tf.MarginBottom=[single]1;$tf.VerticalAnchor=[int]$anchor; $tr=$tf.TextRange;$tr.Text=$text;$tr.Font.Size=[single]$size;$tr.Font.Bold=[int]$bold;$tr.Font.Name=[string]$font;Retry{$tr.Font.Color.RGB=[int](Rgb $color)};$tr.ParagraphFormat.Alignment=[int]$align; try{$tr.ParagraphFormat.SpaceBefore=[single]0;$tr.ParagraphFormat.SpaceAfter=[single]0}catch{}; return $tb }
function Arrow($x1,$y1,$x2,$y2,$color,$w){ $c=$script:slide.Shapes.AddConnector(1,[single]$x1,[single]$y1,[single]$x2,[single]$y2); Retry{$c.Line.ForeColor.RGB=[int](Rgb $color)};$c.Line.Weight=[single]$w;$c.Line.BeginArrowheadStyle=[int]1;$c.Line.EndArrowheadStyle=[int]2;NoShadow $c; return $c }
function DelAll($sl){ foreach($sh in @($sl.Shapes)){ try{$sh.Delete()}catch{} } }
function FindLike($pres,$prefix){ foreach($sl in $pres.Slides){ foreach($sh in $sl.Shapes){ try{ if($sh.HasTextFrame -eq -1 -and $sh.TextFrame.HasText -eq -1){ if($sh.TextFrame.TextRange.Text.Trim().StartsWith($prefix)){ return $sl } } }catch{} } }; return $null }

$s = FindLike $pres "Benchmark absorption"
Write-Output ("s19 idx=" + $s.SlideIndex)
DelAll $s
$script:slide = $s
$bg=$script:slide.Shapes.AddShape(1,[single]0,[single]0,[single]960,[single]540); $bg.Line.Visible=0; NoShadow $bg; $bg.Fill.Solid(); Retry{$bg.Fill.ForeColor.RGB=[int](Rgb $BGD)}
Txt 26 14 760 30 "Benchmark absorption map" 26 -1 $WHITE "Aptos Display" 1 1 | Out-Null
Txt 770 16 158 14 ([string]$s.SlideIndex) 9 0 $MUTE "Aptos" 3 1 | Out-Null
Txt 26 50 760 16 "Absorb each vendor's strength INTO one AVEVA-based, operations-ready PLM" 11 0 $LIGHT "Aptos" 1 1 | Out-Null

# left source cards
$src = @(
 @{v="Siemens — Core PLM rigor"; d="Item · BOM · baseline · effectivity · config & change governance"; c=$SLATE2},
 @{v="Dassault — Experience platform"; d="Role-based collaboration · virtual twin · mfg planning flow"; c=$VIOLET},
 @{v="Hexagon — Asset lifecycle"; d="SDx asset info · completions · j5 operations · project controls"; c=$GREEN},
 @{v="PTC — PLM + ALM thread"; d="Requirements · SW traceability · AI-assisted change (Windchill/Codebeamer)"; c=$ORANGE}
)
$cy=120
$mids=@()
for($i=0;$i -lt 4;$i++){ $r=$src[$i]
  Panel 44 $cy 320 78 $PANEL $r.c 1 0 0.08 | Out-Null
  Bar 44 $cy 5 78 $r.c 0.5 | Out-Null
  Txt 58 ($cy+10) 300 16 $r.v 9.5 -1 $WHITE "Aptos" 1 1 | Out-Null
  Txt 58 ($cy+32) 300 38 $r.d 8 0 $LIGHT "Aptos" 1 1 | Out-Null
  $mids += ($cy+39)
  $cy += 86
}
# hub
Panel 580 150 356 300 "0E2436" $CYAN 1.5 0 0.05 | Out-Null
Bar 580 150 356 6 $CYAN 0 | Out-Null
Txt 596 162 324 18 "New AVEVA-based PLM" 13 -1 $CYAN "Aptos Display" 1 1 | Out-Null
Txt 596 184 324 14 "Engineering -> Operations digital thread" 9 0 $LIGHT "Aptos" 1 1 | Out-Null
Txt 596 208 324 14 "Absorbs the best of each:" 8.5 -1 $MUTE "Aptos" 1 1 | Out-Null
$abs = @("Configuration rigor (Siemens-grade)","Collaboration & virtual twin (Dassault)","Asset lifecycle & completions (Hexagon)","Requirements / ALM + AI change (PTC)")
$ay=228
foreach($a in $abs){ Oval 600 ($ay+4) 8 8 $GREEN | Out-Null; Txt 614 ($ay) 312 14 $a 8.5 0 $LIGHT "Aptos" 1 1 | Out-Null; $ay+=20 }
Bar 596 314 324 1 $GRID 0 | Out-Null
Txt 596 320 324 14 "AVEVA differentiation:" 8.5 -1 $CYAN "Aptos" 1 1 | Out-Null
Txt 596 338 324 36 "E3D/Marine + AIM + PI + UOC + CONNECT  ->  operations-ready, configuration-governed PLM" 8.5 0 $WHITE "Aptos" 1 1 | Out-Null
# arrows converge into hub
for($i=0;$i -lt 4;$i++){ Arrow 364 $mids[$i] 576 (250+$i*22) $src[$i].c 1.75 | Out-Null }
# bottom note
Panel 44 470 892 44 "0E2A24" $GREEN 1 0 0.08 | Out-Null
Txt 56 470 868 44 "Strategy: absorb proven strengths into ONE configuration-governed, executable, operations-ready PLM — not a copy of any single backbone." 9.5 -1 $WHITE "Aptos" 1 3 | Out-Null

$pres.Save()
$s.Export("$outDir\s19_conv.png","PNG",1280,720)
$pres.Close()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($pres) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($ppt)  | Out-Null
[GC]::Collect(); [GC]::WaitForPendingFinalizers()
Write-Output "BUILD8_S19_DONE"
