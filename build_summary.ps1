$ErrorActionPreference = "Stop"
$dir = "D:\이력서\AVEVA - Marine Principal Technical Support & Consultant – PLM SME, Busan\Proposal"
$out = Join-Path $dir "전무님용_핵심요약_슬라이드.pptx"
$png = "C:\Users\namma\.claude\plm_slide_work\review_sum"
New-Item -ItemType Directory -Force -Path $png | Out-Null

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
function Bar($l,$t,$w,$h,$fill,$rad){ $s=$script:slide.Shapes.AddShape(5,[single]$l,[single]$t,[single]$w,[single]$h); try{$s.Adjustments.Item(1)=[single]$rad}catch{}; $s.Fill.Solid(); Retry{$s.Fill.ForeColor.RGB=[int](Rgb $fill)}; $s.Line.Visible=0; NoShadow $s; return $s }
function Oval($l,$t,$w,$h,$fill){ $s=$script:slide.Shapes.AddShape(9,[single]$l,[single]$t,[single]$w,[single]$h); $s.Fill.Solid(); Retry{$s.Fill.ForeColor.RGB=[int](Rgb $fill)}; $s.Line.Visible=0; NoShadow $s; return $s }
function Txt($l,$t,$w,$h,$text,$size,$bold,$color,$font,$align,$anchor){ $tb=$script:slide.Shapes.AddTextbox(1,[single]$l,[single]$t,[single]$w,[single]$h); $tb.Fill.Visible=0;$tb.Line.Visible=0;NoShadow $tb; $tf=$tb.TextFrame;$tf.WordWrap=-1;$tf.AutoSize=0;$tf.MarginLeft=[single]2;$tf.MarginRight=[single]2;$tf.MarginTop=[single]1;$tf.MarginBottom=[single]1;$tf.VerticalAnchor=[int]$anchor; $tr=$tf.TextRange;$tr.Text=$text;$tr.Font.Size=[single]$size;$tr.Font.Bold=[int]$bold;$tr.Font.Name=[string]$font;Retry{$tr.Font.Color.RGB=[int](Rgb $color)};$tr.ParagraphFormat.Alignment=[int]$align; try{$tr.ParagraphFormat.SpaceBefore=[single]0;$tr.ParagraphFormat.SpaceAfter=[single]0}catch{}; return $tb }

$bg=$script:slide.Shapes.AddShape(1,[single]0,[single]0,[single]960,[single]540); $bg.Line.Visible=0; NoShadow $bg; $bg.Fill.Solid(); Retry{$bg.Fill.ForeColor.RGB=[int](Rgb $BGD)}
Txt 26 12 900 28 "AVEVA Marine PLM — HD중공업 수주 전략  ·  1-Page 핵심 요약" 21 -1 $WHITE "Aptos Display" 1 1 | Out-Null
# one-liner banner
Panel 36 48 888 34 "0E2A36" $CYAN 1.25 0.1 | Out-Null
Txt 48 48 70 34 "핵심" 11 -1 $CYAN "Aptos" 1 3 | Out-Null
Txt 122 48 794 34 "지멘스 = 제품 PLM 백본  ·  AVEVA = 설계↔운영 개방형 디지털 스레드 — 고객이 직접 바꿔 쓰는 가볍고 빠른 PLM으로 수주를 이긴다" 11 -1 $WHITE "Aptos" 1 3 | Out-Null
Txt 40 88 884 16 "왜 지금: HD가 지멘스와 저울질 중 · 결정 포인트 = 도입 리스크 · 현업 사용성 · 조선 특화 · 미래 확장성 → 이 네 가지에서 AVEVA 우위" 9 0 $MUTE "Aptos" 1 1 | Out-Null

# left: 지멘스 대비 우위 5
Panel 36 112 504 316 $PANEL "2C4663" 1 0.04 | Out-Null
Bar 36 112 504 4 $CYAN 0 | Out-Null
Txt 50 120 480 18 "지멘스 대비 우위 (5)" 12 -1 $CYAN "Aptos" 1 1 | Out-Null
$adv = @(
 @{c=$CYAN;   t="개방형 API (FastAPI)"; d="HD가 UI·프로세스를 직접 구성·변경 → 벤더 락인 해소"},
 @{c=$BLUE;   t="경량 · 고속"; d="필요한 기능만 호출 → 무겁고 버벅이던 기존 PLM 해소"},
 @{c=$GREEN;  t="엔지니어링 ↔ 운영"; d="E3D/Marine + AIM + PI + CONNECT 디지털 스레드 완성"},
 @{c=$ORANGE; t="AI 품질 게이트"; d="오류·누락·요구사항 미반영 자동 점검 → 무결한 생산 착수"},
 @{c=$VIOLET; t="인프라까지 제시"; d="데이터센터·보안GW·10/25G LACP·이중화 (HW 제안 또는 직접 구축)"}
)
$ry=148; $i=1
foreach($a in $adv){
  Oval 52 ($ry+2) 22 22 $a.c | Out-Null
  Txt 52 ($ry+2) 22 22 ([string]$i) 11 -1 $WHITE "Aptos" 2 3 | Out-Null
  Txt 84 ($ry) 446 16 $a.t 11 -1 $WHITE "Aptos" 1 1 | Out-Null
  Txt 84 ($ry+18) 450 16 $a.d 8.5 0 $LIGHT "Aptos" 1 1 | Out-Null
  $ry += 54; $i++
}

# right: two boxes
Panel 556 112 368 150 $PANEL $GREEN 1 0.05 | Out-Null
Bar 556 112 368 4 $GREEN 0 | Out-Null
Txt 570 120 340 18 "요구사항 → 생산 통제" 12 -1 $GREEN "Aptos" 1 1 | Out-Null
Txt 570 146 344 108 "Requirement 입력 시 VCRM·Check List 자동 생성(관리자 거버넌스). 결과물 입력 + M-BOM/MCO 완료 + 최종 승인 후에만 생산 시작 — 요구사항 누락 없는 납기." 10.5 0 $LIGHT "Aptos" 1 1 | Out-Null

Panel 556 276 368 152 $PANEL $ORANGE 1 0.05 | Out-Null
Bar 556 276 368 4 $ORANGE 0 | Out-Null
Txt 570 284 340 18 "실행 · 관리 (WBS / KPI)" 12 -1 $ORANGE "Aptos" 1 1 | Out-Null
Txt 570 310 344 112 "WBS 0~4 + KPI 게이트(계약 100% · E2E≥95% · BOM/유효성 99% · 영향≥85% · 대시보드≤5s). 관리 주체: Principal Consultant. 대체가 아닌 보완 → 리스크↓ · 조기 가치." 10.5 0 $LIGHT "Aptos" 1 1 | Out-Null

# bottom conclusion banner
Panel 36 438 888 58 "0E2A24" $GREEN 1.25 0.06 | Out-Null
Txt 48 438 90 58 "결론" 12 -1 $GREEN "Aptos" 1 3 | Out-Null
Txt 144 444 770 48 "'있는 것을 파는 제안'이 아니라 '고객 요구를 반영해 더 나은 제품을 빠르게'. 첨부(영문 덱·비교분석·시각자료)로 HD 제안 즉시 보강 가능.   ※ 내부 검토용 아이디어 산출물." 9.5 -1 $WHITE "Aptos" 1 3 | Out-Null

$pres.SaveAs($out, 24)
$pres.Slides.Item(1).Export("$png\summary.png","PNG",1280,720)
$pres.Close()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($pres) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($ppt) | Out-Null
[GC]::Collect(); [GC]::WaitForPendingFinalizers()
Write-Output ("SUMMARY_DONE " + $out)
