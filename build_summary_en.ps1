$ErrorActionPreference = "Stop"
$dir = "D:\이력서\AVEVA - Marine Principal Technical Support & Consultant – PLM SME, Busan\Proposal"
$out = Join-Path $dir "Executive_Summary_1page_EN.pptx"
$png = "C:\Users\namma\.claude\plm_slide_work\review_sum"

function Rgb([string]$hex){ $hex=$hex.TrimStart('#'); [int]([Convert]::ToInt32($hex.Substring(0,2),16)+[Convert]::ToInt32($hex.Substring(2,2),16)*256+[Convert]::ToInt32($hex.Substring(4,2),16)*65536) }
function Retry($sb){ for($a=0;$a -lt 6;$a++){ try{ return (& $sb) }catch{ Start-Sleep -Milliseconds 80 } }; return (& $sb) }
$WHITE="FFFFFF"; $LIGHT="C7D6E5"; $MUTE="93A8C0"; $CYAN="34C0EE"; $BLUE="4F8BE8"; $GREEN="49C28C"; $ORANGE="F0832F"; $VIOLET="9B7BE0"
$BGD="0B1626"; $PANEL="14253B"

$ppt = New-Object -ComObject PowerPoint.Application
$pres = $ppt.Presentations.Add(0)
$pres.PageSetup.SlideWidth=[single]960; $pres.PageSetup.SlideHeight=[single]540
$script:slide = $pres.Slides.Add(1,12)

function NoShadow($s){ try{$s.Shadow.Visible=0}catch{} }
function Panel($l,$t,$w,$h,$fill,$line,$lw,$rad){ $s=$script:slide.Shapes.AddShape(5,[single]$l,[single]$t,[single]$w,[single]$h); try{$s.Adjustments.Item(1)=[single]$rad}catch{}; $s.Fill.Solid(); Retry{$s.Fill.ForeColor.RGB=[int](Rgb $fill)}; if($line){$s.Line.Visible=-1; Retry{$s.Line.ForeColor.RGB=[int](Rgb $line)}; $s.Line.Weight=[single]$lw}else{$s.Line.Visible=0}; NoShadow $s; return $s }
function Bar($l,$t,$w,$h,$fill){ $s=$script:slide.Shapes.AddShape(1,[single]$l,[single]$t,[single]$w,[single]$h); $s.Fill.Solid(); Retry{$s.Fill.ForeColor.RGB=[int](Rgb $fill)}; $s.Line.Visible=0; NoShadow $s; return $s }
function Oval($l,$t,$w,$h,$fill){ $s=$script:slide.Shapes.AddShape(9,[single]$l,[single]$t,[single]$w,[single]$h); $s.Fill.Solid(); Retry{$s.Fill.ForeColor.RGB=[int](Rgb $fill)}; $s.Line.Visible=0; NoShadow $s; return $s }
function Txt($l,$t,$w,$h,$text,$size,$bold,$color,$font,$align,$anchor){ $tb=$script:slide.Shapes.AddTextbox(1,[single]$l,[single]$t,[single]$w,[single]$h); $tb.Fill.Visible=0;$tb.Line.Visible=0;NoShadow $tb; $tf=$tb.TextFrame;$tf.WordWrap=-1;$tf.AutoSize=0;$tf.MarginLeft=[single]2;$tf.MarginRight=[single]2;$tf.MarginTop=[single]1;$tf.MarginBottom=[single]1;$tf.VerticalAnchor=[int]$anchor; $tr=$tf.TextRange;$tr.Text=$text;$tr.Font.Size=[single]$size;$tr.Font.Bold=[int]$bold;$tr.Font.Name=[string]$font;Retry{$tr.Font.Color.RGB=[int](Rgb $color)};$tr.ParagraphFormat.Alignment=[int]$align; try{$tr.ParagraphFormat.SpaceBefore=[single]0;$tr.ParagraphFormat.SpaceAfter=[single]0}catch{}; return $tb }

$bg=$script:slide.Shapes.AddShape(1,[single]0,[single]0,[single]960,[single]540); $bg.Line.Visible=0; NoShadow $bg; $bg.Fill.Solid(); Retry{$bg.Fill.ForeColor.RGB=[int](Rgb $BGD)}
Txt 26 12 900 28 "AVEVA Marine PLM — Strategy to Win HD Hyundai  ·  1-Page Summary" 20 -1 $WHITE "Aptos Display" 1 1 | Out-Null
Panel 36 48 888 34 "0E2A36" $CYAN 1.25 0.1 | Out-Null
Txt 48 48 60 34 "CORE" 11 -1 $CYAN "Aptos" 1 3 | Out-Null
Txt 116 48 800 34 "Siemens = product-PLM backbone · AVEVA = open engineering<->operations digital thread — a light, fast PLM the customer reshapes itself wins the deal." 10.5 -1 $WHITE "Aptos" 1 3 | Out-Null
Txt 40 88 884 16 "Why now: HD is weighing Siemens. Decision = adoption risk · day-to-day usability · shipbuilding fit · future extensibility -> AVEVA wins on all four." 9 0 $MUTE "Aptos" 1 1 | Out-Null

Panel 36 112 504 316 $PANEL "2C4663" 1 0.04 | Out-Null
Bar 36 112 504 4 $CYAN | Out-Null
Txt 50 120 480 18 "Advantages over Siemens (5)" 12 -1 $CYAN "Aptos" 1 1 | Out-Null
$adv = @(
 @{c=$CYAN;   t="Open API (FastAPI)"; d="HD builds/changes its own UI & process -> removes vendor lock-in"},
 @{c=$BLUE;   t="Light & fast"; d="Call only the functions needed -> fixes heavy, stuttering legacy PLM"},
 @{c=$GREEN;  t="Engineering <-> Operations"; d="E3D/Marine + AIM + PI + CONNECT digital thread"},
 @{c=$ORANGE; t="AI quality gate"; d="Auto-checks errors/omissions/unmet requirements -> clean production start"},
 @{c=$VIOLET; t="Full infrastructure"; d="Data center · security GW · 10/25G LACP · redundancy (HW or AVEVA build)"}
)
$ry=148; $i=1
foreach($a in $adv){ Oval 52 ($ry+2) 22 22 $a.c | Out-Null; Txt 52 ($ry+2) 22 22 ([string]$i) 11 -1 $WHITE "Aptos" 2 3 | Out-Null; Txt 84 ($ry) 446 16 $a.t 11 -1 $WHITE "Aptos" 1 1 | Out-Null; Txt 84 ($ry+18) 450 16 $a.d 8.5 0 $LIGHT "Aptos" 1 1 | Out-Null; $ry+=54; $i++ }

Panel 556 112 368 150 $PANEL $GREEN 1 0.05 | Out-Null
Bar 556 112 368 4 $GREEN | Out-Null
Txt 570 120 340 18 "Requirement -> production control" 12 -1 $GREEN "Aptos" 1 1 | Out-Null
Txt 570 146 344 108 "VCRM & Checklist auto-generated on requirement entry (admin governance). Production starts only after deliverables + M-BOM/MCO + final approval — no requirement missed." 10.5 0 $LIGHT "Aptos" 1 1 | Out-Null

Panel 556 276 368 152 $PANEL $ORANGE 1 0.05 | Out-Null
Bar 556 276 368 4 $ORANGE | Out-Null
Txt 570 284 340 18 "Delivery & management (WBS/KPI)" 12 -1 $ORANGE "Aptos" 1 1 | Out-Null
Txt 570 310 344 112 "WBS 0-4 + KPI gates (contract 100% · E2E>=95% · BOM/effectivity 99% · impact>=85% · dashboard<=5s). Owned by the Principal Consultant. Complement, not replace -> lower risk, early value." 10.5 0 $LIGHT "Aptos" 1 1 | Out-Null

Panel 36 438 888 58 "0E2A24" $GREEN 1.25 0.06 | Out-Null
Txt 48 438 130 58 "CONCLUSION" 12 -1 $GREEN "Aptos" 1 3 | Out-Null
Txt 184 444 730 48 "Not selling what already exists — quickly building a better product around the customer's needs. The attached deck / comparison / visuals can strengthen the HD bid immediately.   * Internal review material." 9.5 -1 $WHITE "Aptos" 1 3 | Out-Null

$pres.SaveAs($out, 24)
$pres.Slides.Item(1).Export("$png\summary_en.png","PNG",1280,720)
$pres.Close()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($pres) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($ppt) | Out-Null
[GC]::Collect(); [GC]::WaitForPendingFinalizers()
Write-Output ("SUMMARY_EN_DONE " + $out)
