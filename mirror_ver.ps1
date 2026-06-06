$ErrorActionPreference = "Stop"
$dir = "D:\이력서\AVEVA - Marine Principal Technical Support & Consultant – PLM SME, Busan\Proposal"
$ko  = Join-Path $dir "Future_Industrial_PLM_Meeting_Deck_KO_Rebuilt_v2.pptx"
$koName = "Future_Industrial_PLM_Meeting_Deck_KO_Rebuilt_v2.pptx"
$orig = Join-Path $dir "Future_Industrial_PLM_Meeting_Deck_EN_Rebuilt.pptx"
$origName = "Future_Industrial_PLM_Meeting_Deck_EN_Rebuilt.pptx"
$outDir = "C:\Users\namma\.claude\plm_slide_work\review_ver"

function Rgb([string]$hex){ $hex=$hex.TrimStart('#'); [int]([Convert]::ToInt32($hex.Substring(0,2),16)+[Convert]::ToInt32($hex.Substring(2,2),16)*256+[Convert]::ToInt32($hex.Substring(4,2),16)*65536) }
function Retry($sb){ for($a=0;$a -lt 6;$a++){ try{ return (& $sb) }catch{ Start-Sleep -Milliseconds 80 } }; return (& $sb) }
$WHITE="FFFFFF"; $LIGHT="C7D6E5"; $MUTE="93A8C0"; $CYAN="34C0EE"; $GREEN="49C28C"; $ORANGE="F0832F"; $VIOLET="9B7BE0"; $AMBER="E8B23A"
$BGD="0B1626"; $PANEL="14253B"

$ppt = New-Object -ComObject PowerPoint.Application
function NoShadow($s){ try{$s.Shadow.Visible=0}catch{} }
function Panel($l,$t,$w,$h,$fill,$line,$lw,$rad){ $s=$script:slide.Shapes.AddShape(5,[single]$l,[single]$t,[single]$w,[single]$h); try{$s.Adjustments.Item(1)=[single]$rad}catch{}; $s.Fill.Solid(); Retry{$s.Fill.ForeColor.RGB=[int](Rgb $fill)}; if($line){$s.Line.Visible=-1; Retry{$s.Line.ForeColor.RGB=[int](Rgb $line)}; $s.Line.Weight=[single]$lw}else{$s.Line.Visible=0}; NoShadow $s; return $s }
function Bar($l,$t,$w,$h,$fill,$rad){ $s=$script:slide.Shapes.AddShape(5,[single]$l,[single]$t,[single]$w,[single]$h); try{$s.Adjustments.Item(1)=[single]$rad}catch{}; $s.Fill.Solid(); Retry{$s.Fill.ForeColor.RGB=[int](Rgb $fill)}; $s.Line.Visible=0; NoShadow $s; return $s }
function Oval($l,$t,$w,$h,$fill){ $s=$script:slide.Shapes.AddShape(9,[single]$l,[single]$t,[single]$w,[single]$h); $s.Fill.Solid(); Retry{$s.Fill.ForeColor.RGB=[int](Rgb $fill)}; $s.Line.Visible=0; NoShadow $s; return $s }
function Txt($l,$t,$w,$h,$text,$size,$bold,$color,$font,$align,$anchor){ $tb=$script:slide.Shapes.AddTextbox(1,[single]$l,[single]$t,[single]$w,[single]$h); $tb.Fill.Visible=0;$tb.Line.Visible=0;NoShadow $tb; $tf=$tb.TextFrame;$tf.WordWrap=-1;$tf.AutoSize=0;$tf.MarginLeft=[single]2;$tf.MarginRight=[single]2;$tf.MarginTop=[single]1;$tf.MarginBottom=[single]1;$tf.VerticalAnchor=[int]$anchor; $tr=$tf.TextRange;$tr.Text=$text;$tr.Font.Size=[single]$size;$tr.Font.Bold=[int]$bold;$tr.Font.Name=[string]$font;Retry{$tr.Font.Color.RGB=[int](Rgb $color)};$tr.ParagraphFormat.Alignment=[int]$align; try{$tr.ParagraphFormat.SpaceBefore=[single]0;$tr.ParagraphFormat.SpaceAfter=[single]0}catch{}; return $tb }
function FindLike($pres,$prefix){ foreach($sl in $pres.Slides){ foreach($sh in $sl.Shapes){ try{ if($sh.HasTextFrame -eq -1 -and $sh.TextFrame.HasText -eq -1){ if($sh.TextFrame.TextRange.Text.Trim().StartsWith($prefix)){ return $sl } } }catch{} } }; return $null }

function BuildVer($title,$sub,$chL,$chT,$cards,$coreL,$coreT,$idx){
  $bg=$script:slide.Shapes.AddShape(1,[single]0,[single]0,[single]960,[single]540); $bg.Line.Visible=0; NoShadow $bg; $bg.Fill.Solid(); Retry{$bg.Fill.ForeColor.RGB=[int](Rgb $BGD)}
  Txt 26 12 760 30 $title 21 -1 $WHITE "Aptos Display" 1 1 | Out-Null
  Txt 770 14 158 14 "Future Industrial PLM" 8.5 0 $MUTE "Aptos" 3 1 | Out-Null
  Txt 905 30 30 14 ([string]$idx) 9 0 $MUTE "Aptos" 3 1 | Out-Null
  Txt 26 46 900 16 $sub 11 0 $LIGHT "Aptos" 1 1 | Out-Null
  Panel 40 76 880 30 "31240F" $AMBER 1 0.2 | Out-Null
  Txt 52 76 120 30 $chL 10 -1 $AMBER "Aptos" 1 3 | Out-Null
  Txt 178 76 740 30 $chT 9.5 0 $LIGHT "Aptos" 1 3 | Out-Null
  foreach($k in $cards){ $x=$k.x; $y=$k.y
    Panel $x $y 436 166 $PANEL $k.c 1 0.05 | Out-Null
    Bar $x $y 436 5 $k.c 0 | Out-Null
    Oval ($x+14) ($y+16) 30 30 $k.c | Out-Null
    Txt ($x+14) ($y+16) 30 30 ([string]$k.n) 14 -1 $WHITE "Aptos" 2 3 | Out-Null
    Txt ($x+54) ($y+16) 366 18 $k.en 13 -1 $WHITE "Aptos" 1 1 | Out-Null
    Txt ($x+54) ($y+36) 366 14 $k.sub 9.5 0 $k.c "Aptos" 1 1 | Out-Null
    Txt ($x+16) ($y+62) 406 96 $k.d 11 0 $LIGHT "Aptos" 1 1 | Out-Null
  }
  Panel 40 462 880 48 "0E2A36" $CYAN 1.25 0.06 | Out-Null
  Txt 52 462 96 48 $coreL 11 -1 $CYAN "Aptos" 1 3 | Out-Null
  Txt 156 466 764 44 $coreT 10 -1 $WHITE "Aptos" 1 3 | Out-Null
}

# ===== KO deck (Korean) =====
foreach($p in @($ppt.Presentations)){ if($p.Name -eq $koName){ try{ $p.Saved=$true; $p.Close() }catch{} } }
$pres=$ppt.Presentations.Open($ko,0,0,0)
$an = FindLike $pres "설계툴"
$ins = if($an -ne $null){ [int]($an.SlideIndex+1) } else { [int]($pres.Slides.Count+1) }
$script:slide = $pres.Slides.Add($ins,12)
$koCards=@(
 @{x=40;y=116;n=1;c=$CYAN;en="DB 기반 (파일 아님)";sub="DB element, 통째 파일 아님";d="E3D는 DB 기반입니다. 네이티브 element·extract/merge 변경이력과 PLM 메타데이터·관계만 가볍게 버전관리합니다 — 통째 파일 스냅샷 없이."},
 @{x=484;y=116;n=2;c=$GREEN;en="콘텐츠 주소형 형상 저장";sub="무거운 3D는 한 번만";d="무거운 형상은 오브젝트 스토리지(Git-LFS/S3식·중복제거)에 한 번만 저장하고, PLM은 참조와 델타만 추적해 설계툴·서버가 가볍습니다."},
 @{x=40;y=288;n=3;c=$ORANGE;en="이벤트 기반 스냅샷";sub="promote / baseline 시점";d="매 저장이 아니라 promote/AI 게이트/baseline 승인 시점에 버전을 남깁니다. 이미 설계한 거버넌스 흐름이 곧 버전 트리거입니다."},
 @{x=484;y=288;n=4;c=$VIOLET;en="비동기 큐 캡처";sub="백그라운드, 병목 없음";d="스냅샷은 이중화 리눅스 서버의 큐를 통해 백그라운드로 캡처되어, 다수 설계자가 동시에 작업해도 병목이 없습니다."}
)
BuildVer "E3D / Hull 설계파일 버전관리 방식" "디지털 스레드가 버전관리를 끌고 간다 — 3D 설계를 소스코드처럼 다루지 않는다" "과제" "3D 파일은 방대 — 매 저장 Git식은 과부하 · 중앙 서버는 다중 설계자에 병목 · 수동 업로드는 버전 누락" $koCards "핵심" "E3D는 소스코드가 아니다. DB element + 메타데이터 버전관리 · 콘텐츠 주소형 형상 저장 · promote/baseline 이벤트 기반 비동기 버전 캡처 — 디지털 스레드가 버전관리를 끌고 간다." $ins
$pres.Save(); $pres.Slides.Item($ins).Export("$outDir\ver_ko_mirror.png","PNG",1280,720)
Write-Output ("KO mirror idx=" + $ins); $pres.Close()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($pres) | Out-Null

# ===== Original (English) =====
foreach($p in @($ppt.Presentations)){ if($p.Name -eq $origName){ try{ $p.Saved=$true; $p.Close() }catch{} } }
$pres=$ppt.Presentations.Open($orig,0,0,0)
$an = FindLike $pres "Hardware"
$ins = if($an -ne $null){ [int]($an.SlideIndex+1) } else { [int]($pres.Slides.Count+1) }
$script:slide = $pres.Slides.Add($ins,12)
$enCards=@(
 @{x=40;y=116;n=1;c=$CYAN;en="DB-based, not file-based";sub="database elements, not whole files";d="E3D is database-driven. Version its native element & extract/merge change history plus PLM metadata and relationships - lightweight, with no whole-file snapshots."},
 @{x=484;y=116;n=2;c=$GREEN;en="Content-addressed geometry";sub="store heavy 3D once";d="Heavy geometry is stored once in object storage (Git-LFS / S3 style, de-duplicated). PLM keeps only references and deltas, so the design tool and server stay light."},
 @{x=40;y=288;n=3;c=$ORANGE;en="Event-driven snapshots";sub="on promote / baseline, not every save";d="Version on promote / AI-gate / baseline approval - not on every save. The governance flow already designed becomes the version trigger."},
 @{x=484;y=288;n=4;c=$VIOLET;en="Async queued capture";sub="background, no bottleneck";d="Snapshots are captured in the background via a queue on the redundant Linux servers, so many concurrent designers never hit a bottleneck."}
)
BuildVer "E3D / Hull design - version control approach" "The digital thread drives version control - we do not treat 3D design like source code." "CHALLENGE" "3D files are huge - per-save Git is too heavy  ·  a central server bottlenecks with many concurrent designers  ·  manual upload misses versions" $enCards "CORE" "E3D is not source code. DB-element + metadata versioning · content-addressed geometry · versions triggered by promote/baseline events and captured asynchronously - the digital thread drives version control." $ins
$pres.Save(); $pres.Slides.Item($ins).Export("$outDir\ver_orig_mirror.png","PNG",1280,720)
Write-Output ("ORIG mirror idx=" + $ins); $pres.Close()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($pres) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($ppt) | Out-Null
[GC]::Collect(); [GC]::WaitForPendingFinalizers()
Write-Output "MIRROR_DONE"
