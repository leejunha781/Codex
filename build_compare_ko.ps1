$ErrorActionPreference = "Stop"
$dir = "D:\이력서\AVEVA - Marine Principal Technical Support & Consultant – PLM SME, Busan\Proposal"
$ko = Join-Path $dir "Future_Industrial_PLM_Meeting_Deck_KO_Rebuilt_v2.pptx"
$out = Join-Path $dir "AVEVA_비교분석_Comparison_KO.pptx"
$outName = "AVEVA_비교분석_Comparison_KO.pptx"
$png = "C:\Users\namma\.claude\plm_slide_work\review_cmpko"
New-Item -ItemType Directory -Force -Path $png | Out-Null
Copy-Item -LiteralPath $ko -Destination $out -Force

function Rgb([string]$hex){ $hex=$hex.TrimStart('#'); [int]([Convert]::ToInt32($hex.Substring(0,2),16)+[Convert]::ToInt32($hex.Substring(2,2),16)*256+[Convert]::ToInt32($hex.Substring(4,2),16)*65536) }
function Retry($sb){ for($a=0;$a -lt 6;$a++){ try{ return (& $sb) }catch{ Start-Sleep -Milliseconds 80 } }; return (& $sb) }
$WHITE="FFFFFF"; $LIGHT="C7D6E5"; $MUTE="93A8C0"; $CYAN="34C0EE"; $GREEN="49C28C"; $AMBER="E8B23A"; $GREY="8FA9C4"
$BGD="0B1626"; $PANEL="14253B"; $ROW="101F33"; $HL="0E2A38"

$ppt = New-Object -ComObject PowerPoint.Application
foreach($p in @($ppt.Presentations)){ if($p.Name -eq $outName){ try{ $p.Saved=$true; $p.Close() }catch{} } }
$pres=$ppt.Presentations.Open($out,0,0,0)

function NoShadow($s){ try{$s.Shadow.Visible=0}catch{} }
function Panel($l,$t,$w,$h,$fill,$line,$lw,$rad){ $s=$script:slide.Shapes.AddShape(5,[single]$l,[single]$t,[single]$w,[single]$h); try{$s.Adjustments.Item(1)=[single]$rad}catch{}; $s.Fill.Solid(); Retry{$s.Fill.ForeColor.RGB=[int](Rgb $fill)}; if($line){$s.Line.Visible=-1; Retry{$s.Line.ForeColor.RGB=[int](Rgb $line)}; $s.Line.Weight=[single]$lw}else{$s.Line.Visible=0}; NoShadow $s; return $s }
function Bar($l,$t,$w,$h,$fill){ $s=$script:slide.Shapes.AddShape(1,[single]$l,[single]$t,[single]$w,[single]$h); $s.Fill.Solid(); Retry{$s.Fill.ForeColor.RGB=[int](Rgb $fill)}; $s.Line.Visible=0; NoShadow $s; return $s }
function Oval($l,$t,$w,$h,$fill){ $s=$script:slide.Shapes.AddShape(9,[single]$l,[single]$t,[single]$w,[single]$h); $s.Fill.Solid(); Retry{$s.Fill.ForeColor.RGB=[int](Rgb $fill)}; $s.Line.Visible=0; NoShadow $s; return $s }
function Txt($l,$t,$w,$h,$text,$size,$bold,$color,$font,$align,$anchor){ $tb=$script:slide.Shapes.AddTextbox(1,[single]$l,[single]$t,[single]$w,[single]$h); $tb.Fill.Visible=0;$tb.Line.Visible=0;NoShadow $tb; $tf=$tb.TextFrame;$tf.WordWrap=-1;$tf.AutoSize=0;$tf.MarginLeft=[single]2;$tf.MarginRight=[single]2;$tf.MarginTop=[single]1;$tf.MarginBottom=[single]1;$tf.VerticalAnchor=[int]$anchor; $tr=$tf.TextRange;$tr.Text=$text;$tr.Font.Size=[single]$size;$tr.Font.Bold=[int]$bold;$tr.Font.Name=[string]$font;Retry{$tr.Font.Color.RGB=[int](Rgb $color)};$tr.ParagraphFormat.Alignment=[int]$align; try{$tr.ParagraphFormat.SpaceBefore=[single]0;$tr.ParagraphFormat.SpaceAfter=[single]0}catch{}; return $tb }
function HasTitle($sl,$titles){ foreach($sh in $sl.Shapes){ try{ if($sh.HasTextFrame -eq -1 -and $sh.TextFrame.HasText -eq -1){ $t=$sh.TextFrame.TextRange.Text.Trim(); foreach($k in $titles){ if($t -eq $k){ return $true } } } }catch{} }; return $false }

$keep = @("현재 PLM 지형","공개 로드맵 시그널","채워야 할 전략적 공백","벤치마크 흡수 맵","근거 자료")
for($i=$pres.Slides.Count; $i -ge 1; $i--){ $sl=$pres.Slides.Item($i); if(-not (HasTitle $sl $keep)){ $sl.Delete() } }
Write-Output ("kept(KO) = " + $pres.Slides.Count)

# cover
$script:slide = $pres.Slides.Add(1,12)
$bg=$script:slide.Shapes.AddShape(1,[single]0,[single]0,[single]960,[single]540); $bg.Line.Visible=0; NoShadow $bg; $bg.Fill.Solid(); Retry{$bg.Fill.ForeColor.RGB=[int](Rgb $BGD)}
Bar 60 250 120 5 $CYAN | Out-Null
Txt 60 150 840 22 "BENCHMARK / 비교분석" 13 -1 $CYAN "Aptos" 1 1 | Out-Null
Txt 58 180 850 60 "AVEVA vs 타 PLM · 조선 설계툴 — 비교분석" 30 -1 $WHITE "Aptos Display" 1 1 | Out-Null
Txt 60 270 850 22 "현재 상태 + 로드맵(Roadmap) 기준  ·  HD중공업 수주 검토용" 14 0 $LIGHT "Aptos" 1 1 | Out-Null
Txt 60 320 850 20 "①현재 PLM 지형  ②공개 로드맵 시그널  ③전략적 공백  ④벤치마크 흡수맵  ⑤근거 자료  ⑥조선 설계툴 비교" 11 0 $MUTE "Aptos" 1 1 | Out-Null
Txt 60 470 850 16 "※ 공개 출처 기반 내부 검토용 · Siemens / Dassault / Hexagon / PTC vs AVEVA" 9 0 $MUTE "Aptos" 1 1 | Out-Null

# shipbuilding tool matrix (Korean) at end
$ins = [int]($pres.Slides.Count+1)
$script:slide = $pres.Slides.Add($ins,12)
$bg=$script:slide.Shapes.AddShape(1,[single]0,[single]0,[single]960,[single]540); $bg.Line.Visible=0; NoShadow $bg; $bg.Fill.Solid(); Retry{$bg.Fill.ForeColor.RGB=[int](Rgb $BGD)}
Txt 26 12 900 28 "조선 설계툴 비교 — AVEVA vs 경쟁 (현재 + 로드맵)" 20 -1 $WHITE "Aptos Display" 1 1 | Out-Null
Txt 26 46 900 16 "AVEVA E3D/Marine만이 선체+의장+생산 설계와 운영·PLM 통합을 함께 제공" 11 0 $LIGHT "Aptos" 1 1 | Out-Null
$colX=@(214,356,498,640,782); $colW=142; $dimX=36; $dimW=174
$tools=@("AVEVA`nE3D/Marine","SENER`nFORAN","CADMATIC","NAPA","Hexagon`nSmart 3D")
Panel $colX[0] 118 $colW 340 $HL $CYAN 1.25 0.03 | Out-Null
for($c=0;$c -lt 5;$c++){ $col = if($c -eq 0){$CYAN}else{$WHITE}; Txt $colX[$c] 120 $colW 34 ($tools[$c]) 10 -1 $col "Aptos" 2 3 | Out-Null }
Txt $dimX 120 $dimW 34 "역량" 9.5 -1 $MUTE "Aptos" 1 3 | Out-Null
$dims=@("선체 (Hull)","의장 (Outfitting)","생산 (Production)","복원성·초기설계","PLM·운영 통합","클라우드·AI 로드맵")
$mat=@(@(1,1,2,2,2),@(1,1,1,3,1),@(1,1,2,3,2),@(2,2,3,1,3),@(1,3,2,3,2),@(1,2,2,2,2))
$sym=@{1="●";2="◐";3="○"}; $word=@{1="강함";2="부분";3="제한"}; $col=@{1=$GREEN;2=$AMBER;3=$GREY}
$ry=158; $rh=49
for($r=0;$r -lt 6;$r++){
  if($r % 2 -eq 1){ Panel $dimX $ry (910-$dimX) ($rh-3) $ROW $null 0 0.02 | Out-Null }
  Txt $dimX ($ry) $dimW ($rh-3) $dims[$r] 9.5 -1 $LIGHT "Aptos" 1 3 | Out-Null
  for($c=0;$c -lt 5;$c++){ $lv=$mat[$r][$c]; Txt $colX[$c] ($ry) $colW ($rh-3) ($sym[$lv]+"  "+$word[$lv]) 9.5 -1 $col[$lv] "Aptos" 2 3 | Out-Null }
  $ry += $rh
}
Panel 36 460 888 42 "0E2A24" $GREEN 1.25 0.06 | Out-Null
Txt 48 460 90 42 "시사점" 10 -1 $GREEN "Aptos" 1 3 | Out-Null
Txt 148 462 768 38 "AVEVA만이 선체·의장·생산 설계와 운영·PLM(AIM/PI/CONNECT)을 함께 제공 → 설계-운영 디지털 스레드. NAPA(복원성)·CADMATIC(의장)은 강점이 한정적." 9.5 -1 $WHITE "Aptos" 1 3 | Out-Null
Txt 36 506 888 14 "● 강함   ◐ 부분   ○ 제한      ※ 공개 정보 기반 일반 포지셔닝 · 버전/구성에 따라 상이" 8 0 $MUTE "Aptos" 1 1 | Out-Null

$pres.SaveAs($out, 24)
for($i=1;$i -le $pres.Slides.Count;$i++){ $num="{0:D2}" -f $i; $pres.Slides.Item($i).Export("$png\cmpko-$num.png","PNG",1280,720) }
Write-Output ("FINAL KO compare = " + $pres.Slides.Count)
$pres.Close()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($pres) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($ppt) | Out-Null
[GC]::Collect(); [GC]::WaitForPendingFinalizers()
Write-Output ("COMPARE_KO_DONE " + $out)
