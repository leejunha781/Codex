# ==== Genohco portfolio builder REV2 (Genohco past-experience removed; focus 대양전기/인텔리안) ====
$ErrorActionPreference = 'Stop'
trap { Write-Output ("TRAP @line " + $_.InvocationInfo.ScriptLineNumber + " :: " + ($_.InvocationInfo.Line).Trim() + " :: " + $_.Exception.Message); break }
function C([int]$r,[int]$g,[int]$b){ return ($r + ($g*256) + ($b*65536)) }

$navy  = C 18 31 56
$navy2 = C 27 43 74
$steel = C 44 82 130
$cyan  = C 0 183 200
$amber = C 245 166 35
$bg    = C 244 246 249
$white = C 255 255 255
$ink   = C 23 30 44
$muted = C 96 106 122
$line  = C 223 228 235
$chip  = C 234 240 247
$green = C 21 138 90
$F = '맑은 고딕'

$ppt = New-Object -ComObject PowerPoint.Application
$ppt.Visible = 1
$pres = $ppt.Presentations.Add(1)
$pres.PageSetup.SlideWidth  = 960
$pres.PageSetup.SlideHeight = 540

function NewSlide(){ return $pres.Slides.Add($pres.Slides.Count+1, 12) }
function Rect($sl,$l,$t,$w,$h,$fill){ $s=$sl.Shapes.AddShape(1,$l,$t,$w,$h); $s.Fill.Solid(); $s.Fill.ForeColor.RGB=$fill; $s.Line.Visible=0; $s.Shadow.Visible=0; return $s }
function Bord($sl,$l,$t,$w,$h,$fill,$lc,$lw){ $s=$sl.Shapes.AddShape(1,$l,$t,$w,$h); $s.Fill.Solid(); $s.Fill.ForeColor.RGB=$fill; $s.Line.Visible=-1; $s.Line.ForeColor.RGB=$lc; $s.Line.Weight=$lw; $s.Shadow.Visible=0; return $s }
function RR($sl,$l,$t,$w,$h,$fill,$rad){ $s=$sl.Shapes.AddShape(5,$l,$t,$w,$h); $s.Fill.Solid(); $s.Fill.ForeColor.RGB=$fill; $s.Line.Visible=0; $s.Shadow.Visible=0; try{$s.Adjustments.Item(1)=$rad}catch{}; return $s }
function Txt($sl,$l,$t,$w,$h,$text,$size,$bold,$col,$align,$anchor){
  $tb=$sl.Shapes.AddTextbox(1,$l,$t,$w,$h)
  $tf=$tb.TextFrame; $tf.WordWrap=-1; $tf.AutoSize=0
  $tf.MarginLeft=0;$tf.MarginRight=0;$tf.MarginTop=0;$tf.MarginBottom=0
  if($anchor){ $tf.VerticalAnchor=[int]$anchor }
  $tr=$tf.TextRange; $tr.Text=$text
  $tr.Font.Size=[single]$size
  $tr.Font.Bold=[int]$bold
  $tr.Font.Color.RGB=[int]$col
  $tr.Font.Name=[string]$F
  if($align){ $tr.ParagraphFormat.Alignment=[int]$align }
  return $tb }
function Head($sl,$title,$kicker,$num){
  Rect $sl 0 0 960 540 $bg | Out-Null
  Rect $sl 54 46 6 30 $cyan | Out-Null
  Txt $sl 70 42 640 38 $title 25 -1 $ink 1 1 | Out-Null
  if($kicker){ Txt $sl 70 80 760 22 $kicker 12.5 0 $muted 1 1 | Out-Null }
  Rect $sl 54 116 852 1.4 $line | Out-Null
  Txt $sl 820 44 86 20 ("0"+$num+" / 08") 11 0 $muted 3 1 | Out-Null
  Txt $sl 612 80 294 20 "GENOHCO · 시스템 체계" 11 0 $cyan 3 1 | Out-Null }

# ---------- SLIDE 1 : COVER ----------
$s = NewSlide
Rect $s 0 0 960 540 $navy | Out-Null
Rect $s 0 0 960 6 $cyan | Out-Null
$g1 = RR $s 560 -80 520 520 $navy2 0.06; $g1.Rotation=18
Rect $s 80 70 6 26 $cyan | Out-Null
Txt $s 96 68 700 24 "SYSTEM ENGINEERING PORTFOLIO" 13 -1 $cyan 1 1 | Out-Null
Txt $s 80 150 700 80 "이준하" 56 -1 $white 1 1 | Out-Null
Txt $s 82 232 780 34 "위성·항공 시스템 체계 엔지니어  |  Senior System Engineer" 21 -1 $cyan 1 1 | Out-Null
Txt $s 82 276 800 26 "방산 통신·C4I 체계통합  →  LEO 위성통신 단말·ESA 안테나 통합검증      |      총 경력 21년 7개월" 13 0 (C 174 186 205) 1 1 | Out-Null
$tl = RR $s 80 330 700 70 $navy2 0.10
Txt $s 100 330 660 70 "요구사항을 아키텍처·링크 버짓·시험 evidence로 전환해, 검증된 시스템으로 완성합니다." 15 -1 $white 1 3 | Out-Null
Rect $s 80 446 700 1 (C 70 84 110) | Out-Null
Txt $s 80 458 760 24 "leejunha781@gmail.com        ·        010-2731-4581        ·        경기 군포시" 13 0 (C 174 186 205) 1 1 | Out-Null
Txt $s 80 486 760 22 "㈜제노코 시스템기술연구소 · 시스템 체계 직무 지원" 12.5 -1 $amber 1 1 | Out-Null

# ---------- SLIDE 2 : 핵심 요약 ----------
$s = NewSlide
Head $s "한눈에 보는 핵심" "21년 7개월의 시스템 체계 역량을 수치와 키워드로 요약" 2
$kpi = @(
 @('21.7년','시스템 체계 경력', $navy),
 @('30% ↓','월간 불량률 감소', $cyan),
 @('25% ↓','고객 이슈 리드타임', $steel),
 @('400+','설계검증 체크리스트', $navy),
 @('약 12억','신규 사업 수주 기여(₩)', $amber)
)
$kx=54; $kw=160; $kg=13; $ky=140; $kh=118
for($i=0;$i -lt $kpi.Count;$i++){
  $x=$kx + $i*($kw+$kg)
  Bord $s $x $ky $kw $kh $white $line 1 | Out-Null
  Rect $s $x $ky $kw 5 $kpi[$i][2] | Out-Null
  Txt $s $x ($ky+24) $kw 46 $kpi[$i][0] 30 -1 $kpi[$i][2] 2 1 | Out-Null
  Txt $s $x ($ky+82) $kw 22 $kpi[$i][1] 12 0 $muted 2 1 | Out-Null
}
Bord $s 54 286 852 92 $white $line 1 | Out-Null
Rect $s 54 286 5 92 $cyan | Out-Null
Txt $s 74 296 812 76 "방산 통신체계와 LEO 위성통신 단말·ESA 안테나 분야에서 시스템 설계·체계종합, 조립·통합 절차, 환경시험, 아키텍처·링크 버짓 분석을 직접 수행해 온 시스템 체계 엔지니어입니다. 상위 요구사항을 인터페이스·시험조건·검증 evidence로 전환하고, RF/HW/SW/네트워크 성능을 Link/Power Budget 기준으로 통합검증합니다." 13.5 0 $ink 1 3 | Out-Null
Txt $s 54 392 300 22 "핵심 역량 키워드" 12.5 -1 $steel 1 1 | Out-Null
$tags = @('방산 통신·C4I 체계통합','시스템 설계·체계종합','Link/Power Budget','아키텍처 분석','조립·통합 절차','LEO 위성통신 단말','Ku-band ESA','EMI/EMC·진동','RF/안테나','DFE·ADC/DAC','통합검증','FAT/HAT/SAT','방산 산출물','고객·시험기관 대응')
$cx=54; $cyy=418; $maxR=906; $chh=30
foreach($it in $tags){
  $tw=0; foreach($ch in $it.ToCharArray()){ if([int][char]$ch -gt 127){$tw+=12.6}else{$tw+=6.4} }
  $cw=[int]($tw+22)
  if($cx+$cw -gt $maxR){ $cx=54; $cyy+=$chh+8 }
  RR $s $cx $cyy $cw $chh $chip 0.5 | Out-Null
  Txt $s $cx $cyy $cw $chh $it 12 0 $steel 2 3 | Out-Null
  $cx+=$cw+8
}

# ---------- SLIDE 3 : 경력 여정 (제노코 past-experience 제외) ----------
$s = NewSlide
Head $s "경력 여정 — 회로에서 시스템 체계까지" "방산 체계통합(14.9년)과 LEO 위성통신 상용검증으로 완성한 시스템 체계 역량" 3
$tlY=300
Rect $s 110 $tlY 740 3 $steel | Out-Null
$steps = @(
 @('2004–2007','초기 R&D','회로·보드·RF bring-up','네스랩·에세텔·엑시큐어하이트론', $steel),
 @('2007–2022','방산 통신·C4I PM/PL · 14.9년','KSS-III·FFX Batch-II·C4I/LINK-11','대양전기공업', $navy),
 @('2023–2025','LEO 단말·Ku-band ESA 통합검증','품질·field issue·양산 release readiness','인텔리안테크놀로지스 · 부장', $cyan),
 @('2026 ~','시스템 체계 직무 지원','위성체 설계·체계종합·우주환경시험','㈜제노코', $amber)
)
$n=$steps.Count; $cw=176; $gap=(740-($n*$cw))/($n-1); $x0=110
for($i=0;$i -lt $n;$i++){
  $cx=$x0 + $i*($cw+$gap)
  $nodeX = $cx + $cw/2
  $nd=$s.Shapes.AddShape(9, ($nodeX-9), ($tlY-7.5), 18, 18); $nd.Fill.Solid(); $nd.Fill.ForeColor.RGB=$steps[$i][4]; $nd.Line.Visible=-1; $nd.Line.ForeColor.RGB=$white; $nd.Line.Weight=2.5; $nd.Shadow.Visible=0
  $above = ($i % 2 -eq 0)
  if($above){ $cardY=$tlY-150 } else { $cardY=$tlY+34 }
  Bord $s $cx $cardY $cw 122 $white $line 1 | Out-Null
  Rect $s $cx $cardY $cw 5 $steps[$i][4] | Out-Null
  Txt $s ($cx+12) ($cardY+12) ($cw-24) 22 $steps[$i][0] 13 -1 $steps[$i][4] 1 1 | Out-Null
  Txt $s ($cx+12) ($cardY+36) ($cw-24) 38 $steps[$i][1] 12 -1 $ink 1 1 | Out-Null
  Txt $s ($cx+12) ($cardY+74) ($cw-24) 30 $steps[$i][2] 10.3 0 $muted 1 1 | Out-Null
  Txt $s ($cx+12) ($cardY+102) ($cw-24) 16 $steps[$i][3] 10 -1 $steel 1 1 | Out-Null
  if($above){ Rect $s ($nodeX-0.75) ($cardY+122) 1.5 ($tlY-($cardY+122)) $line | Out-Null }
  else { Rect $s ($nodeX-0.75) ($tlY+3) 1.5 ($cardY-($tlY+3)) $line | Out-Null }
}

# ---------- SLIDE 4 : 역량 맵 ----------
$s = NewSlide
Head $s "역량 맵 — 시스템 체계 5대 도메인" "JD가 요구하는 5개 영역을 모두 실무로 커버" 4
$bars = @(
 @('방산 통신 · C4I 체계통합 (PM/PL)', 95, '시스템 아키텍처 · ICD · FAT/HAT/SAT · 고객 인수 · 양산이관', $navy),
 @('아키텍처 · Link/Power Budget 분석', 93, 'EIRP · G/T · NF · C/N0 · scan loss · pointing', $cyan),
 @('LEO 위성통신 단말 · ESA 안테나 통합검증', 90, 'RF/안테나 · DFE · ADC/DAC · Ethernet · SW log', $steel),
 @('조립 · 통합 · 환경시험', 86, 'FAT/HAT/SAT · EMI/EMC · 진동 · 방수 (열진공 요구사항 검토)', $navy),
 @('방산 산출물 · 고객/시험기관 대응', 93, '제안서 · WBS · 시험계획/절차/성적서 · 검사기관', $steel)
)
$by=146; $bh=66; $bgap=8; $trackX=360; $trackW=470
for($i=0;$i -lt $bars.Count;$i++){
  $y=$by + $i*($bh+$bgap)
  Bord $s 54 $y 852 $bh $white $line 1 | Out-Null
  Txt $s 70 ($y+11) 290 22 $bars[$i][0] 13.5 -1 $ink 1 1 | Out-Null
  Txt $s 70 ($y+36) 290 20 $bars[$i][2] 10 0 $muted 1 1 | Out-Null
  RR $s $trackX ($y+24) $trackW 14 (C 232 236 242) 0.5 | Out-Null
  $fw=[int]($trackW * $bars[$i][1] / 100)
  RR $s $trackX ($y+24) $fw 14 $bars[$i][3] 0.5 | Out-Null
  Txt $s ($trackX+$trackW+8) ($y+20) 60 22 ([string]$bars[$i][1]+'%') 13 -1 $bars[$i][3] 1 1 | Out-Null
}

# ---------- SLIDE 5 : JD 적합성 (제노코 명칭 제외) ----------
$s = NewSlide
Head $s "제노코 시스템 체계 JD 적합성" "채용 요건과 보유 경험을 1:1로 매핑 — 빈칸 없는 적합성" 5
$rows = @(
 @('학력 · 경력 · 전공','석사(임베디드 시스템) · 21년 7개월 · 전자/제어계측 — 필수요건 충족'),
 @('위성체 시스템 설계 · 체계종합','인텔리안 LEO 단말·Ku-band ESA 체계종합 + 대양전기 방산 체계통합 PM/PL'),
 @('시스템 조립 / 통합 절차','통합검증 · FAT/HAT/SAT · 양산 release readiness 전 주기 수행'),
 @('우주 환경시험(열진공·EMI/EMC)','EMI/EMC·진동 환경시험 + 열진공(TVAC) 요구사항 검토·시험기관 대응'),
 @('아키텍처 · 링크 버짓 분석','Link/Power Budget로 EIRP·G/T·NF·C/N0 통합검증 (핵심강점)'),
 @('우주·항공 전장품 개발 경력자','방산 통신체계 · LEO 위성통신 · ESA 안테나 21년 — 즉시 전력화')
)
$ry=144; $rh=56; $rgap=6
for($i=0;$i -lt $rows.Count;$i++){
  $y=$ry + $i*($rh+$rgap)
  Bord $s 54 $y 852 $rh $white $line 1 | Out-Null
  Rect $s 54 $y 5 $rh $cyan | Out-Null
  $ck=$s.Shapes.AddShape(9, 72, ($y+($rh/2)-11), 22, 22); $ck.Fill.Solid(); $ck.Fill.ForeColor.RGB=$green; $ck.Line.Visible=0; $ck.Shadow.Visible=0
  Txt $s 72 ($y+($rh/2)-11) 22 22 "✓" 12 -1 $white 2 3 | Out-Null
  Txt $s 108 $y 300 $rh $rows[$i][0] 13.5 -1 $navy 1 3 | Out-Null
  $ar=$s.Shapes.AddShape(1, 410, ($y+($rh/2)-1), 18, 2); $ar.Fill.Solid(); $ar.Fill.ForeColor.RGB=$muted; $ar.Line.Visible=0; $ar.Shadow.Visible=0
  Txt $s 440 $y 452 $rh $rows[$i][1] 12.5 0 $ink 1 3 | Out-Null
}

# ---------- SLIDE 6 : 대표 프로젝트 (대양전기 3 + 인텔리안 1) ----------
$s = NewSlide
Head $s "대표 프로젝트" "방산 체계통합과 LEO 위성통신 상용검증에서 만든 정량 성과" 6
$proj = @(
 @('Eutelsat OneWeb LEO 단말 · Ku-band ESA 통합검증','인텔리안테크놀로지스 · 부장/Senior Engineer', @('RF/안테나·DFE·ADC/DAC·네트워크·SW log·환경시험 통합 원인분석','400+ 설계검증 체크리스트 기반 반복 검증 체계 운영'),'불량률 30%↓ · 리드타임 25%↓', $cyan),
 @('KSS-III 잠수함 통합통신체계','대양전기공업 · PM/PL', @('백본망 + 유·무선 Ethernet 통합 네트워크 구축','ICD·FAT/HAT/SAT·항해시험·고객 인수 전 주기 수행'),'함정 3척 개발 · 인수 완료', $navy),
 @('FFX Batch-II 함내 무선통신·방송체계','대양전기공업 · PM/PL', @('VHF/UHF/HF 출력·수신감도·주파수 정렬·채널 안정성 계측 검증','시스템 통합·시험계획·고객 인수시험 수행'),'수상함 3척 통신체계 개발', $steel),
 @('인도네시아 잠수함 무선 네트워크 아키텍처','대양전기공업 · 제안/수주', @('기존 방식 대신 무선 네트워크 구조를 창의적으로 제안','고객사·조선소 기술 설득 + 개발비·원가 산정'),'약 12억 원 수주 기여', $amber)
)
$px=@(54,486); $py=@(140,300); $cw=420; $ch=148
for($i=0;$i -lt 4;$i++){
  $x=$px[$i % 2]; $y=$py[[int][Math]::Floor($i/2)]
  Bord $s $x $y $cw $ch $white $line 1 | Out-Null
  Rect $s $x $y 5 $ch $proj[$i][4] | Out-Null
  Txt $s ($x+18) ($y+12) ($cw-34) 40 $proj[$i][0] 13.5 -1 $ink 1 1 | Out-Null
  Txt $s ($x+18) ($y+50) ($cw-34) 18 $proj[$i][1] 10.5 -1 $proj[$i][4] 1 1 | Out-Null
  $b = ($proj[$i][2] | ForEach-Object { "•  " + $_ }) -join "`r"
  Txt $s ($x+18) ($y+72) ($cw-34) 44 $b 10.8 0 $muted 1 1 | Out-Null
  RR $s ($x+18) ($y+$ch-32) ($cw-36) 22 $proj[$i][4] 0.5 | Out-Null
  Txt $s ($x+18) ($y+$ch-32) ($cw-36) 22 $proj[$i][3] 11.5 -1 $white 2 3 | Out-Null
}

# ---------- SLIDE 7 : 제니언 인재상 ----------
$s = NewSlide
Head $s "제니언 인재상 적합성" "인재 · 기술 · 품질 · 효율 — 제노코가 강조하는 네 가지 가치에 부합" 7
$jen = @(
 @('인재를 만드는 제니언','직접 최대 7명·8명 TFT 리딩, 후배 엔지니어의 시스템 관점·검증 기준 성장 지원, HW/SW/RF/기구/품질/생산·협력사·고객 간 언어 정합', $navy),
 @('기술을 기본으로 하는 제니언','회로~RF~네트워크~환경시험 풀스택, IBM Machine Learning·Python 및 IELTS로 기술 폭을 스스로 보완', $cyan),
 @('품질을 중시하는 제니언','defect log·CAPA·ECO/BOM·재검증·evidence package를 끝까지 추적, 일정에 쫓겨도 품질 근거와 비타협', $steel),
 @('효율을 극대화하는 제니언','요구사항→아키텍처·ICD·산출물 구조화로 재작업 감소, 무선 네트워크 아키텍처 창의적 제안으로 약 12억 수주', $amber)
)
$qx=@(54,486); $qy=@(142,300); $cw=420; $ch=146
for($i=0;$i -lt 4;$i++){
  $x=$qx[$i % 2]; $y=$qy[[int][Math]::Floor($i/2)]
  Bord $s $x $y $cw $ch $white $line 1 | Out-Null
  RR $s ($x+18) ($y+18) 40 40 $jen[$i][2] 0.25 | Out-Null
  Txt $s ($x+18) ($y+18) 40 40 ([string]($i+1)) 18 -1 $white 2 3 | Out-Null
  Txt $s ($x+70) ($y+20) ($cw-86) 26 $jen[$i][0] 14.5 -1 $ink 1 1 | Out-Null
  Txt $s ($x+70) ($y+50) ($cw-90) 86 $jen[$i][1] 11.8 0 $muted 1 1 | Out-Null
}

# ---------- SLIDE 8 : 90일 플랜 + CLOSING ----------
$s = NewSlide
Rect $s 0 0 960 540 $navy | Out-Null
Rect $s 0 0 960 6 $cyan | Out-Null
Rect $s 54 50 6 30 $cyan | Out-Null
Txt $s 70 46 700 34 "입사 후 90일 실행 플랜" 25 -1 $white 1 1 | Out-Null
Txt $s 70 84 760 22 "재적응이 아니라, 첫 분기부터 검증된 산출물로 기여합니다" 12.5 0 (C 174 186 205) 1 1 | Out-Null
$ph = @(
 @('0–30일','담당 과제 요구사항·인터페이스·시험일정·산출물·리스크 파악 / 추적표·ICD·절차서·시험문서 일관성 점검', $cyan),
 @('31–60일','Link/Power Budget·환경시험(TVAC·EMI/EMC) 요구사항을 설계검토·시험 readiness 체크리스트로 구체화 / field issue 종결 체계 정착', $steel),
 @('61–90일','체계종합 산출물 완성도 향상 / 시험·검사기관·고객 대응 기술근거 선제 정리 → 일정·품질·납기 리스크 축소', (C 90 160 210)),
 @('90일 +','위성탑재체·지상국·EGSE·항공전자·방산 전장품 검증 산출물 신뢰성 향상 / KAI 항공우주·방산 성장방향 연계', $amber)
)
$y0=128; $rh=62; $rg=10
for($i=0;$i -lt 4;$i++){
  $y=$y0+$i*($rh+$rg)
  RR $s 54 $y 120 $rh $ph[$i][2] 0.16 | Out-Null
  Txt $s 54 $y 120 $rh $ph[$i][0] 16 -1 $white 2 3 | Out-Null
  RR $s 184 $y 722 $rh $navy2 0.1 | Out-Null
  Txt $s 204 $y 686 $rh $ph[$i][1] 12.5 0 (C 208 216 228) 1 3 | Out-Null
}
Rect $s 54 426 852 1 (C 70 84 110) | Out-Null
Txt $s 54 440 852 44 "방산 통신체계 14년과 LEO 위성통신·ESA 안테나 상용검증 경험을 제노코 시스템 체계에 그대로 이식해, 첫 분기부터 검증된 산출물로 기여할 준비가 된 지원자입니다." 14 -1 $white 2 3 | Out-Null
Txt $s 54 500 852 22 "이준하  ·  leejunha781@gmail.com  ·  010-2731-4581" 12 0 (C 174 186 205) 2 1 | Out-Null

# ---------- SAVE ----------
$tmpx = "$env:TEMP\genohco_portfolio_rev2.pptx"
$tmpd = "$env:TEMP\genohco_portfolio_rev2.pdf"
try{ [System.IO.File]::Delete($tmpx) }catch{}
try{ [System.IO.File]::Delete($tmpd) }catch{}
$pres.SaveAs($tmpx, 24)
$pres.SaveAs($tmpd, 32)
$pres.Close()
$ppt.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($pres) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($ppt) | Out-Null
Write-Output ("BUILT rev2 pptx=" + (Get-Item $tmpx).Length + "B  pdf=" + (Get-Item $tmpd).Length + "B")
