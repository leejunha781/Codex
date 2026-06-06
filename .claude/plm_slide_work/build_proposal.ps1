$ErrorActionPreference = "Stop"
$dir = "D:\이력서\AVEVA - Marine Principal Technical Support & Consultant – PLM SME, Busan\Proposal"
$out = Join-Path $dir "HD현대중공업_제안서_양식.pptx"
$png = "C:\Users\namma\.claude\plm_slide_work\review_prop"
New-Item -ItemType Directory -Force -Path $png | Out-Null

function Rgb([string]$hex){ $hex=$hex.TrimStart('#'); [int]([Convert]::ToInt32($hex.Substring(0,2),16)+[Convert]::ToInt32($hex.Substring(2,2),16)*256+[Convert]::ToInt32($hex.Substring(4,2),16)*65536) }
function Retry($sb){ for($a=0;$a -lt 6;$a++){ try{ return (& $sb) }catch{ Start-Sleep -Milliseconds 80 } }; return (& $sb) }
$WHITE="FFFFFF"; $INK="1E2D38"; $GRAY="6B7C88"; $HD="00A887"; $HDD="0B7A66"; $LINE="D8E0E4"; $LIGHT="EEF4F4"; $TINT="DCF1EC"

$ppt = New-Object -ComObject PowerPoint.Application
$pres = $ppt.Presentations.Add(0)   # windowless
$pres.PageSetup.SlideWidth = [single]960
$pres.PageSetup.SlideHeight = [single]540

function NoShadow($s){ try{$s.Shadow.Visible=0}catch{} }
function Panel($l,$t,$w,$h,$fill,$line,$lw,$rad){ $s=$script:slide.Shapes.AddShape(5,[single]$l,[single]$t,[single]$w,[single]$h); try{$s.Adjustments.Item(1)=[single]$rad}catch{}; $s.Fill.Solid(); Retry{$s.Fill.ForeColor.RGB=[int](Rgb $fill)}; if($line){$s.Line.Visible=-1; Retry{$s.Line.ForeColor.RGB=[int](Rgb $line)}; $s.Line.Weight=[single]$lw}else{$s.Line.Visible=0}; NoShadow $s; return $s }
function Rect($l,$t,$w,$h,$fill){ $s=$script:slide.Shapes.AddShape(1,[single]$l,[single]$t,[single]$w,[single]$h); $s.Fill.Solid(); Retry{$s.Fill.ForeColor.RGB=[int](Rgb $fill)}; $s.Line.Visible=0; NoShadow $s; return $s }
function RectDash($l,$t,$w,$h,$line){ $s=$script:slide.Shapes.AddShape(1,[single]$l,[single]$t,[single]$w,[single]$h); $s.Fill.Visible=0; $s.Line.Visible=-1; Retry{$s.Line.ForeColor.RGB=[int](Rgb $line)}; $s.Line.Weight=[single]1; try{$s.Line.DashStyle=4}catch{}; NoShadow $s; return $s }
function Txt($l,$t,$w,$h,$text,$size,$bold,$color,$font,$align,$anchor){ $tb=$script:slide.Shapes.AddTextbox(1,[single]$l,[single]$t,[single]$w,[single]$h); $tb.Fill.Visible=0;$tb.Line.Visible=0;NoShadow $tb; $tf=$tb.TextFrame;$tf.WordWrap=-1;$tf.AutoSize=0;$tf.MarginLeft=[single]2;$tf.MarginRight=[single]2;$tf.MarginTop=[single]1;$tf.MarginBottom=[single]1;$tf.VerticalAnchor=[int]$anchor; $tr=$tf.TextRange;$tr.Text=$text;$tr.Font.Size=[single]$size;$tr.Font.Bold=[int]$bold;$tr.Font.Name=[string]$font;Retry{$tr.Font.Color.RGB=[int](Rgb $color)};$tr.ParagraphFormat.Alignment=[int]$align; try{$tr.ParagraphFormat.SpaceBefore=[single]0;$tr.ParagraphFormat.SpaceAfter=[single]0}catch{}; return $tb }
function Line2($x1,$y1,$x2,$y2,$color,$w){ $c=$script:slide.Shapes.AddConnector(1,[single]$x1,[single]$y1,[single]$x2,[single]$y2); Retry{$c.Line.ForeColor.RGB=[int](Rgb $color)};$c.Line.Weight=[single]$w;$c.Line.BeginArrowheadStyle=[int]1;$c.Line.EndArrowheadStyle=[int]1;NoShadow $c; return $c }
function Chev($l,$t,$w,$h,$fill,$trans){ $tp=52; try{ $s=$script:slide.Shapes.AddShape($tp,[single]$l,[single]$t,[single]$w,[single]$h) }catch{ $s=$script:slide.Shapes.AddShape(7,[single]$l,[single]$t,[single]$w,[single]$h); $s.Rotation=[single]90 }; $s.Fill.Solid(); Retry{$s.Fill.ForeColor.RGB=[int](Rgb $fill)}; $s.Fill.Transparency=[single]$trans; $s.Line.Visible=0; NoShadow $s; return $s }

function Logo($x,$y){
  $c=Rect $x ($y+1) 30 20 $HD; $c.Fill.Solid(); Retry{$c.Fill.ForeColor.RGB=[int](Rgb $HD)}
  $tf=$c.TextFrame; $tf.VerticalAnchor=[int]3; $tr=$tf.TextRange; $tr.Text="HD"; $tr.Font.Size=[single]11; $tr.Font.Bold=[int](-1); $tr.Font.Name="Arial"; Retry{$tr.Font.Color.RGB=[int](Rgb $WHITE)}; $tr.ParagraphFormat.Alignment=[int]2
  Txt ($x+34) $y 90 22 "현대중공업" 12 -1 $INK "Aptos" 1 3 | Out-Null
}
function Header($pageStr){
  Rect 0 0 960 540 $WHITE | Out-Null
  Txt 40 16 520 14 "Unlocking the Limitless Potential of the Ocean" 9 0 $GRAY "Aptos" 1 1 | Out-Null
  Logo 800 14
  Line2 40 510 920 510 $LINE 1 | Out-Null
  Txt 40 514 500 14 "HD현대중공업 제안서  ·  CONFIDENTIAL" 8 0 $GRAY "Aptos" 1 1 | Out-Null
  Txt 740 514 180 14 $pageStr 8 0 $GRAY "Aptos" 3 1 | Out-Null
}
function TitleBar($title){
  Rect 40 54 10 24 $HD | Out-Null
  Txt 58 52 700 28 $title 20 -1 $INK "Aptos Display" 1 3 | Out-Null
  Line2 40 86 920 86 $LINE 1 | Out-Null
}

# ===== Slide 1: COVER =====
$script:slide = $pres.Slides.Add(1,12)
Rect 0 0 960 540 $WHITE | Out-Null
# right chevron motif
Chev 720 70 150 400 $TINT 0.2 | Out-Null
Chev 800 70 150 400 $HD 0.86 | Out-Null
Txt 40 30 520 16 "Unlocking the Limitless Potential of the Ocean" 9 0 $GRAY "Aptos" 1 1 | Out-Null
Logo 800 26
Txt 44 150 400 18 "P R O P O S A L   ·   제 안 서" 12 -1 $HD "Aptos" 1 1 | Out-Null
Rect 44 174 130 4 $HD | Out-Null
Txt 44 190 700 84 "[ 제안 제목을 입력하세요 ]" 31 -1 $INK "Aptos Display" 1 1 | Out-Null
Txt 46 276 700 24 "[ 부제 · 제안 범위를 입력하세요 ]" 14 0 $GRAY "Aptos" 1 1 | Out-Null
Panel 44 350 540 132 $LIGHT $LINE 1 0.04 | Out-Null
$meta = @("제출처 (To)         :   HD현대중공업","제안사 (From)       :   [ 회사 / 팀명 ]","담당자 (Contact)    :   [ 성명 · 직위 · 연락처 ]","제출일 (Date)       :   [ YYYY. MM. DD ]")
$my=364
foreach($m in $meta){ Txt 64 $my 500 22 $m 12 0 $INK "Consolas" 1 1 | Out-Null; $my+=28 }
Txt 40 512 880 16 "HD현대중공업 제안서  ·  CONFIDENTIAL" 8 0 $GRAY "Aptos" 1 1 | Out-Null

# ===== Slide 2: 목차 =====
$script:slide = $pres.Slides.Add(2,12)
Header "02 / 05"
TitleBar "목차   Contents"
$toc = @(
 @{n="01"; k="제안 개요"; e="Executive Summary"},
 @{n="02"; k="현황 및 과제"; e="Current Status & Challenges"},
 @{n="03"; k="제안 솔루션"; e="Proposed Solution"},
 @{n="04"; k="추진 계획 및 일정 (WBS)"; e="Plan & Schedule"},
 @{n="05"; k="기대효과 및 결론"; e="Expected Benefits & Conclusion"}
)
$ty=120
foreach($it in $toc){
  Txt 60 $ty 70 40 $it.n 26 -1 $HD "Aptos Display" 1 3 | Out-Null
  Txt 140 ($ty+2) 700 22 $it.k 16 -1 $INK "Aptos" 1 1 | Out-Null
  Txt 140 ($ty+24) 700 16 $it.e 10 0 $GRAY "Aptos" 1 1 | Out-Null
  Line2 140 ($ty+46) 900 ($ty+46) $LINE 0.75 | Out-Null
  $ty += 70
}

# ===== Slide 3: 본문 양식 1 =====
$script:slide = $pres.Slides.Add(3,12)
Header "03 / 05"
TitleBar "[ 섹션 제목 ]"
Panel 40 100 880 40 $TINT $HD 1 0.2 | Out-Null
Txt 54 100 120 40 "핵심 메시지" 11 -1 $HDD "Aptos" 1 3 | Out-Null
Txt 170 100 740 40 "[ 한 줄 요약 메시지를 입력하세요 ]" 12 0 $INK "Aptos" 1 3 | Out-Null
$by=160
foreach($b in @("[ 내용을 입력하세요 ]","[ 내용을 입력하세요 ]","[ 내용을 입력하세요 ]")){
  Rect 56 ($by+6) 8 8 $HD | Out-Null
  Txt 74 $by 560 24 $b 13 0 $INK "Aptos" 1 1 | Out-Null
  $by += 40
}
Panel 660 160 260 280 $LIGHT $LINE 1 0.04 | Out-Null
Txt 676 172 230 18 "핵심 수치 · 포인트" 11 -1 $HDD "Aptos" 1 1 | Out-Null
Txt 676 200 230 230 "[ 핵심 수치, KPI,¶도식 또는 표를¶입력하세요 ]" 11 0 $GRAY "Aptos" 1 1 | Out-Null

# ===== Slide 4: 본문 양식 2 (텍스트 + 이미지) =====
$script:slide = $pres.Slides.Add(4,12)
Header "04 / 05"
TitleBar "[ 섹션 제목 ]"
Txt 40 104 420 18 "주요 내용" 12 -1 $HDD "Aptos" 1 1 | Out-Null
$by=132
foreach($b in @("[ 내용 1 ]","[ 내용 2 ]","[ 내용 3 ]","[ 내용 4 ]")){
  Rect 48 ($by+6) 8 8 $HD | Out-Null
  Txt 66 $by 400 24 $b 12 0 $INK "Aptos" 1 1 | Out-Null
  $by += 38
}
RectDash 488 104 432 300 $HD | Out-Null
Txt 488 230 432 40 "[ 이미지 · 도식 · 표 영역 ]" 13 0 $GRAY "Aptos" 2 1 | Out-Null
Panel 40 420 880 48 $TINT $HD 1 0.1 | Out-Null
Txt 54 420 120 48 "기대효과" 11 -1 $HDD "Aptos" 1 3 | Out-Null
Txt 170 420 740 48 "[ 본 섹션의 기대효과 / 핵심 포인트를 한 줄로 입력하세요 ]" 12 0 $INK "Aptos" 1 3 | Out-Null

# ===== Slide 5: 맺음 =====
$script:slide = $pres.Slides.Add(5,12)
Rect 0 0 960 540 $WHITE | Out-Null
Chev 740 70 150 400 $HD 0.86 | Out-Null
Txt 40 30 520 16 "Unlocking the Limitless Potential of the Ocean" 9 0 $GRAY "Aptos" 1 1 | Out-Null
Logo 800 26
Rect 60 230 110 5 $HD | Out-Null
Txt 60 250 700 56 "감사합니다" 40 -1 $INK "Aptos Display" 1 1 | Out-Null
Txt 62 312 700 24 "Thank You" 16 0 $GRAY "Aptos" 1 1 | Out-Null
Txt 62 360 700 20 "문의 :  [ 담당자 · 직위 · 이메일 · 연락처 ]" 12 0 $INK "Aptos" 1 1 | Out-Null
Txt 40 512 880 16 "HD현대중공업 제안서  ·  CONFIDENTIAL" 8 0 $GRAY "Aptos" 1 1 | Out-Null

$pres.SaveAs($out, 24)
for($i=1;$i -le 5;$i++){ $pres.Slides.Item($i).Export("$png\prop-$i.png","PNG",1280,720) }
$pres.Close()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($pres) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($ppt) | Out-Null
[GC]::Collect(); [GC]::WaitForPendingFinalizers()
Write-Output ("PROPOSAL_DONE " + $out)
