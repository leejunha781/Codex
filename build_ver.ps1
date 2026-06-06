$ErrorActionPreference = "Stop"
$dir = "D:\이력서\AVEVA - Marine Principal Technical Support & Consultant – PLM SME, Busan\Proposal"
$v2 = Join-Path $dir "Future_Industrial_PLM_Meeting_Deck_EN_Rebuilt_v2.pptx"
$name = "Future_Industrial_PLM_Meeting_Deck_EN_Rebuilt_v2.pptx"
$outDir = "C:\Users\namma\.claude\plm_slide_work\review_ver"
New-Item -ItemType Directory -Force -Path $outDir | Out-Null

function Rgb([string]$hex){ $hex=$hex.TrimStart('#'); [int]([Convert]::ToInt32($hex.Substring(0,2),16)+[Convert]::ToInt32($hex.Substring(2,2),16)*256+[Convert]::ToInt32($hex.Substring(4,2),16)*65536) }
function Retry($sb){ for($a=0;$a -lt 6;$a++){ try{ return (& $sb) }catch{ Start-Sleep -Milliseconds 80 } }; return (& $sb) }
$WHITE="FFFFFF"; $LIGHT="C7D6E5"; $MUTE="93A8C0"; $CYAN="34C0EE"; $GREEN="49C28C"; $ORANGE="F0832F"; $VIOLET="9B7BE0"; $AMBER="E8B23A"
$BGD="0B1626"; $PANEL="14253B"; $LINEC="2C4663"

$ppt = New-Object -ComObject PowerPoint.Application
foreach($p in @($ppt.Presentations)){ if($p.Name -eq $name){ try{ $p.Saved=$true; $p.Close() }catch{} } }
$pres=$ppt.Presentations.Open($v2,0,0,0)

function NoShadow($s){ try{$s.Shadow.Visible=0}catch{} }
function Panel($l,$t,$w,$h,$fill,$line,$lw,$rad){ $s=$script:slide.Shapes.AddShape(5,[single]$l,[single]$t,[single]$w,[single]$h); try{$s.Adjustments.Item(1)=[single]$rad}catch{}; $s.Fill.Solid(); Retry{$s.Fill.ForeColor.RGB=[int](Rgb $fill)}; if($line){$s.Line.Visible=-1; Retry{$s.Line.ForeColor.RGB=[int](Rgb $line)}; $s.Line.Weight=[single]$lw}else{$s.Line.Visible=0}; NoShadow $s; return $s }
function Bar($l,$t,$w,$h,$fill,$rad){ $s=$script:slide.Shapes.AddShape(5,[single]$l,[single]$t,[single]$w,[single]$h); try{$s.Adjustments.Item(1)=[single]$rad}catch{}; $s.Fill.Solid(); Retry{$s.Fill.ForeColor.RGB=[int](Rgb $fill)}; $s.Line.Visible=0; NoShadow $s; return $s }
function Oval($l,$t,$w,$h,$fill){ $s=$script:slide.Shapes.AddShape(9,[single]$l,[single]$t,[single]$w,[single]$h); $s.Fill.Solid(); Retry{$s.Fill.ForeColor.RGB=[int](Rgb $fill)}; $s.Line.Visible=0; NoShadow $s; return $s }
function Txt($l,$t,$w,$h,$text,$size,$bold,$color,$font,$align,$anchor){ $tb=$script:slide.Shapes.AddTextbox(1,[single]$l,[single]$t,[single]$w,[single]$h); $tb.Fill.Visible=0;$tb.Line.Visible=0;NoShadow $tb; $tf=$tb.TextFrame;$tf.WordWrap=-1;$tf.AutoSize=0;$tf.MarginLeft=[single]2;$tf.MarginRight=[single]2;$tf.MarginTop=[single]1;$tf.MarginBottom=[single]1;$tf.VerticalAnchor=[int]$anchor; $tr=$tf.TextRange;$tr.Text=$text;$tr.Font.Size=[single]$size;$tr.Font.Bold=[int]$bold;$tr.Font.Name=[string]$font;Retry{$tr.Font.Color.RGB=[int](Rgb $color)};$tr.ParagraphFormat.Alignment=[int]$align; try{$tr.ParagraphFormat.SpaceBefore=[single]0;$tr.ParagraphFormat.SpaceAfter=[single]0}catch{}; return $tb }
function FindLike($pres,$prefix){ foreach($sl in $pres.Slides){ foreach($sh in $sl.Shapes){ try{ if($sh.HasTextFrame -eq -1 -and $sh.TextFrame.HasText -eq -1){ if($sh.TextFrame.TextRange.Text.Trim().StartsWith($prefix)){ return $sl } } }catch{} } }; return $null }

$anchorSlide = FindLike $pres "Design tool"
$ins = if($anchorSlide -ne $null){ [int]($anchorSlide.SlideIndex+1) } else { [int]($pres.Slides.Count+1) }
$script:slide = $pres.Slides.Add($ins, 12)
Write-Output ("NEW VER SLIDE idx=" + $ins)

$bg=$script:slide.Shapes.AddShape(1,[single]0,[single]0,[single]960,[single]540); $bg.Line.Visible=0; NoShadow $bg; $bg.Fill.Solid(); Retry{$bg.Fill.ForeColor.RGB=[int](Rgb $BGD)}
Txt 26 12 760 30 "E3D / Hull design — version control approach" 22 -1 $WHITE "Aptos Display" 1 1 | Out-Null
Txt 770 14 158 14 "Future Industrial PLM" 8.5 0 $MUTE "Aptos" 3 1 | Out-Null
Txt 905 30 30 14 ([string]$ins) 9 0 $MUTE "Aptos" 3 1 | Out-Null
Txt 26 46 900 16 "디지털 스레드가 버전관리를 끌고 간다  ·  3D 설계를 소스코드처럼 다루지 않는다" 11 0 $LIGHT "Aptos" 1 1 | Out-Null

# challenge strip
Panel 40 76 880 30 "31240F" $AMBER 1 0.2 | Out-Null
Txt 52 76 120 30 "과제 Challenge" 10 -1 $AMBER "Aptos" 1 3 | Out-Null
Txt 178 76 736 30 "3D 파일은 방대 — 매 저장 Git식은 과부하 · 중앙 서버는 다중 설계자에 병목 · 수동 업로드는 버전 누락" 9.5 0 $LIGHT "Aptos" 1 3 | Out-Null

$cards = @(
 @{x=40;  y=116; n=1; c=$CYAN;   en="DB-based, not file-based"; kr="DB 기반 (파일 아님)"; d="E3D is database-driven. Version its native element & extract/merge change history + PLM metadata and relationships — lightweight, no whole-file snapshots."},
 @{x=484; y=116; n=2; c=$GREEN;  en="Content-addressed geometry"; kr="콘텐츠 주소형 형상 저장"; d="Heavy 3D stored once in object storage (Git-LFS / S3 style, de-duplicated). PLM keeps only references + deltas, so the design tool and server stay light."},
 @{x=40;  y=288; n=3; c=$ORANGE; en="Event-driven snapshots"; kr="이벤트 기반 스냅샷"; d="Version on promote / AI-gate / baseline approval — not on every save. The governance flow you already designed becomes the version trigger."},
 @{x=484; y=288; n=4; c=$VIOLET; en="Async queued capture"; kr="비동기 큐 캡처"; d="Snapshots are captured in the background via a queue on the redundant Linux servers, so many designers never hit a bottleneck."}
)
foreach($k in $cards){ $x=$k.x; $y=$k.y
  Panel $x $y 436 166 $PANEL $k.c 1 0.05 | Out-Null
  Bar $x $y 436 5 $k.c 0 | Out-Null
  Oval ($x+14) ($y+16) 30 30 $k.c | Out-Null
  Txt ($x+14) ($y+16) 30 30 ([string]$k.n) 14 -1 $WHITE "Aptos" 2 3 | Out-Null
  Txt ($x+54) ($y+16) 360 18 $k.en 13 -1 $WHITE "Aptos" 1 1 | Out-Null
  Txt ($x+54) ($y+36) 360 14 $k.kr 9.5 0 $k.c "Aptos" 1 1 | Out-Null
  Txt ($x+16) ($y+62) 406 96 $k.d 11 0 $LIGHT "Aptos" 1 1 | Out-Null
}

# core principle banner
Panel 40 462 880 48 "0E2A36" $CYAN 1.25 0.06 | Out-Null
Txt 52 462 120 48 "핵심 Core" 11 -1 $CYAN "Aptos" 1 3 | Out-Null
Txt 178 466 742 44 "E3D is NOT source code.  DB-element + metadata versioning · content-addressed geometry · versions triggered by promote/baseline events, captured async — the digital thread drives version control." 10 -1 $WHITE "Aptos" 1 3 | Out-Null

$pres.Save()
$pres.Slides.Item($ins).Export("$outDir\ver.png","PNG",1280,720)
$pres.Close()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($pres) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($ppt) | Out-Null
[GC]::Collect(); [GC]::WaitForPendingFinalizers()
Write-Output ("VER_DONE idx=" + $ins)
