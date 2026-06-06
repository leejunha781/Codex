$ErrorActionPreference = "Stop"
$dir = "D:\이력서\AVEVA - Marine Principal Technical Support & Consultant – PLM SME, Busan\Proposal"
$cmp = Join-Path $dir "AVEVA_비교분석_Comparison.pptx"
$cmpName = "AVEVA_비교분석_Comparison.pptx"
$png = "C:\Users\namma\.claude\plm_slide_work\review_cmp"

function Rgb([string]$hex){ $hex=$hex.TrimStart('#'); [int]([Convert]::ToInt32($hex.Substring(0,2),16)+[Convert]::ToInt32($hex.Substring(2,2),16)*256+[Convert]::ToInt32($hex.Substring(4,2),16)*65536) }
function Retry($sb){ for($a=0;$a -lt 6;$a++){ try{ return (& $sb) }catch{ Start-Sleep -Milliseconds 80 } }; return (& $sb) }
$WHITE="FFFFFF"; $LIGHT="C7D6E5"; $MUTE="93A8C0"; $CYAN="34C0EE"; $GREEN="49C28C"; $AMBER="E8B23A"; $GREY="8FA9C4"
$BGD="0B1626"; $PANEL="14253B"; $ROW="101F33"; $HL="0E2A38"

$ppt = New-Object -ComObject PowerPoint.Application
foreach($p in @($ppt.Presentations)){ if($p.Name -eq $cmpName){ try{ $p.Saved=$true; $p.Close() }catch{} } }
$pres=$ppt.Presentations.Open($cmp,0,0,0)
$ins = [int]($pres.Slides.Count+1)
$script:slide = $pres.Slides.Add($ins,12)

function NoShadow($s){ try{$s.Shadow.Visible=0}catch{} }
function Panel($l,$t,$w,$h,$fill,$line,$lw,$rad){ $s=$script:slide.Shapes.AddShape(5,[single]$l,[single]$t,[single]$w,[single]$h); try{$s.Adjustments.Item(1)=[single]$rad}catch{}; $s.Fill.Solid(); Retry{$s.Fill.ForeColor.RGB=[int](Rgb $fill)}; if($line){$s.Line.Visible=-1; Retry{$s.Line.ForeColor.RGB=[int](Rgb $line)}; $s.Line.Weight=[single]$lw}else{$s.Line.Visible=0}; NoShadow $s; return $s }
function Bar($l,$t,$w,$h,$fill){ $s=$script:slide.Shapes.AddShape(1,[single]$l,[single]$t,[single]$w,[single]$h); $s.Fill.Solid(); Retry{$s.Fill.ForeColor.RGB=[int](Rgb $fill)}; $s.Line.Visible=0; NoShadow $s; return $s }
function Txt($l,$t,$w,$h,$text,$size,$bold,$color,$font,$align,$anchor){ $tb=$script:slide.Shapes.AddTextbox(1,[single]$l,[single]$t,[single]$w,[single]$h); $tb.Fill.Visible=0;$tb.Line.Visible=0;NoShadow $tb; $tf=$tb.TextFrame;$tf.WordWrap=-1;$tf.AutoSize=0;$tf.MarginLeft=[single]2;$tf.MarginRight=[single]2;$tf.MarginTop=[single]1;$tf.MarginBottom=[single]1;$tf.VerticalAnchor=[int]$anchor; $tr=$tf.TextRange;$tr.Text=$text;$tr.Font.Size=[single]$size;$tr.Font.Bold=[int]$bold;$tr.Font.Name=[string]$font;Retry{$tr.Font.Color.RGB=[int](Rgb $color)};$tr.ParagraphFormat.Alignment=[int]$align; try{$tr.ParagraphFormat.SpaceBefore=[single]0;$tr.ParagraphFormat.SpaceAfter=[single]0}catch{}; return $tb }

$bg=$script:slide.Shapes.AddShape(1,[single]0,[single]0,[single]960,[single]540); $bg.Line.Visible=0; NoShadow $bg; $bg.Fill.Solid(); Retry{$bg.Fill.ForeColor.RGB=[int](Rgb $BGD)}
Txt 26 12 900 28 "Shipbuilding design tools - AVEVA vs peers (current + roadmap)" 20 -1 $WHITE "Aptos Display" 1 1 | Out-Null
Txt 26 46 900 16 "AVEVA E3D/Marine uniquely spans hull + outfitting + production AND operations/PLM integration" 11 0 $LIGHT "Aptos" 1 1 | Out-Null

$colX = @(214,356,498,640,782); $colW=142; $dimX=36; $dimW=174
$tools = @("AVEVA`nE3D/Marine","SENER`nFORAN","CADMATIC","NAPA","Hexagon`nSmart 3D")
# AVEVA column highlight
Panel $colX[0] 118 $colW 340 $HL $CYAN 1.25 0.03 | Out-Null
# header
for($c=0;$c -lt 5;$c++){ $col = if($c -eq 0){$CYAN}else{$WHITE}; Txt $colX[$c] 120 $colW 34 ($tools[$c]) 10 -1 $col "Aptos" 2 3 | Out-Null }
Txt $dimX 120 $dimW 34 "Capability" 9.5 -1 $MUTE "Aptos" 1 3 | Out-Null

$dims = @("Hull","Outfitting","Production","Naval arch / stability","PLM / operations integration","Cloud / AI roadmap")
$mat = @(
 @(1,1,2,2,2),
 @(1,1,1,3,1),
 @(1,1,2,3,2),
 @(2,2,3,1,3),
 @(1,3,2,3,2),
 @(1,2,2,2,2)
)
$sym=@{1="●";2="◐";3="○"}; $word=@{1="Strong";2="Partial";3="Limited"}; $col=@{1=$GREEN;2=$AMBER;3=$GREY}
$ry=158; $rh=49
for($r=0;$r -lt 6;$r++){
  if($r % 2 -eq 1){ Panel $dimX $ry (910-$dimX) ($rh-3) $ROW $null 0 0.02 | Out-Null }
  Txt $dimX ($ry) $dimW ($rh-3) $dims[$r] 9.5 -1 $LIGHT "Aptos" 1 3 | Out-Null
  for($c=0;$c -lt 5;$c++){ $lv=$mat[$r][$c]; Txt $colX[$c] ($ry) $colW ($rh-3) ($sym[$lv]+"  "+$word[$lv]) 9.5 -1 $col[$lv] "Aptos" 2 3 | Out-Null }
  $ry += $rh
}

# takeaway + legend + disclaimer
Panel 36 460 888 42 "0E2A24" $GREEN 1.25 0.06 | Out-Null
Txt 48 460 110 42 "TAKEAWAY" 10 -1 $GREEN "Aptos" 1 3 | Out-Null
Txt 168 462 748 38 "Only AVEVA combines hull/outfitting/production design WITH operations & PLM (AIM/PI/CONNECT) - a true design-to-operations digital thread. NAPA (stability) and CADMATIC (outfitting) are strong but niche." 9.5 -1 $WHITE "Aptos" 1 3 | Out-Null
Txt 36 506 888 14 "● Strong   ◐ Partial   ○ Limited      ※ General positioning from public information; varies by version/configuration." 8 0 $MUTE "Aptos" 1 1 | Out-Null

$pres.Save()
$pres.Slides.Item($ins).Export("$png\cmp-ship.png","PNG",1280,720)
$pres.Close()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($pres) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($ppt) | Out-Null
[GC]::Collect(); [GC]::WaitForPendingFinalizers()
Write-Output ("SHIPTOOLS_EN_DONE idx=" + $ins)
