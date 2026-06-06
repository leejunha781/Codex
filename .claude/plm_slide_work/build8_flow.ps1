$ErrorActionPreference = "Stop"
$v2 = "D:\이력서\AVEVA - Marine Principal Technical Support & Consultant – PLM SME, Busan\Proposal\Future_Industrial_PLM_Meeting_Deck_EN_Rebuilt_v2.pptx"
$name = "Future_Industrial_PLM_Meeting_Deck_EN_Rebuilt_v2.pptx"
$outDir = "C:\Users\namma\.claude\plm_slide_work\review_png8"
New-Item -ItemType Directory -Force -Path $outDir | Out-Null

function Rgb([string]$hex){ $hex=$hex.TrimStart('#'); [int]([Convert]::ToInt32($hex.Substring(0,2),16)+[Convert]::ToInt32($hex.Substring(2,2),16)*256+[Convert]::ToInt32($hex.Substring(4,2),16)*65536) }
function Retry($sb){ for($a=0;$a -lt 6;$a++){ try{ return (& $sb) }catch{ Start-Sleep -Milliseconds 80 } }; return (& $sb) }
$WHITE="FFFFFF"; $LIGHT="C7D6E5"; $MUTE="93A8C0"
$CYAN="34C0EE"; $BLUE="4F8BE8"; $GREEN="49C28C"; $ORANGE="F0832F"; $VIOLET="9B7BE0"; $MAG="D957B8"; $GOLD="E8C24A"
$BGD="0B1626"; $PANEL="14253C"; $BARC="12273C"; $HEXBG="241038"
$TAGA_BG="0E2E3A"; $TAGA_TX="56C8E8"; $TAGP_BG="231A3D"; $TAGP_TX="B79BE6"

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
function DelAll($sl){ foreach($sh in @($sl.Shapes)){ try{$sh.Delete()}catch{} } }
function FindLike($pres,$prefix){ foreach($sl in $pres.Slides){ foreach($sh in $sl.Shapes){ try{ if($sh.HasTextFrame -eq -1 -and $sh.TextFrame.HasText -eq -1){ if($sh.TextFrame.TextRange.Text.Trim().StartsWith($prefix)){ return $sl } } }catch{} } }; return $null }

$flow = FindLike $pres "AI-gated"
Write-Output ("flow idx=" + $flow.SlideIndex)
DelAll $flow
$script:slide = $flow

Bar 0 0 960 540 $BGD 0 | Out-Null
Txt 26 10 760 26 "AI-gated AVEVA design -> PLM flow  (HD Hyundai shipbuilding, 결재선 기반)" 18 -1 $WHITE "Aptos Display" 1 1 | Out-Null
Txt 770 12 158 14 "Future Industrial PLM" 8.5 0 $MUTE "Aptos" 3 1 | Out-Null
Txt 905 28 30 14 ([string]$flow.SlideIndex) 9 0 $MUTE "Aptos" 3 1 | Out-Null
Txt 26 38 908 16 "상세설계(E3D) -> 생산도면(Draw/DWG) -> 자재(MTO) -> 결재 -> 유효성(호선·블록) -> 기준선 -> 설계변경 -> 생산BOM -> 영향분석·승인.  Each step AI-checked, role-gated." 9 0 $LIGHT "Aptos" 1 1 | Out-Null

Bar 22 60 916 22 $BARC 0.25 | Out-Null
Txt 34 60 600 22 "Future PLM   -   API workspace   ·   /api/v1   ·   live status tracked" 9 0 $MUTE "Consolas" 1 3 | Out-Null
Oval 854 67 8 8 $GREEN | Out-Null
Txt 866 60 70 22 "RUNNING" 8.5 -1 $GREEN "Aptos" 1 3 | Out-Null
Txt 36 86 880 16 "AVEVA 설계(E3D/Marine + Draw)에서 생성 -> API PLM에서 거버넌스·추적, 모든 단계 AI 게이트 + 역할(권한) 통제" 10 -1 $CYAN "Aptos" 1 1 | Out-Null

$cards = @(
 @{x=40;  y=112; n=1; kr="상세설계"; en="Detail Design"; tag="AVEVA E3D/Marine"; ta=1; d="호선·블록 3D 모델 / hull & outfitting"; st="modeled"; c=$CYAN},
 @{x=360; y=112; n=2; kr="생산도면"; en="2D Production Drawing"; tag="AVEVA Draw"; ta=1; d="2D sheet -> AutoCAD DWG / DXF"; st="DWG issued"; c=$CYAN},
 @{x=680; y=112; n=3; kr="자재산출"; en="E-BOM / MTO"; tag="AVEVA MTO"; ta=1; d="모델 기반 자재산출 (MTO)"; st="extracted"; c=$CYAN},
 @{x=680; y=246; n=4; kr="결재·승인"; en="Approval"; tag="PLM · 결재선"; ta=0; d="담당 -> 검토 -> 승인 (Eng->Check->Approve)"; st="approved"; c=$BLUE},
 @{x=360; y=246; n=5; kr="유효성"; en="Effectivity"; tag="PLM · API"; ta=0; d="호선 Hull / 블록 Block / Zone / Date"; st="scoped"; c=$GREEN},
 @{x=40;  y=246; n=6; kr="형상기준선"; en="Baseline"; tag="PLM · API"; ta=0; d="Frozen configuration / released rev"; st="baselined"; c=$GREEN},
 @{x=40;  y=380; n=7; kr="설계변경"; en="ECR -> ECO"; tag="PLM · API"; ta=0; d="설계변경요청 -> 변경통보 (change order)"; st="ECO closed"; c=$ORANGE},
 @{x=360; y=380; n=8; kr="생산BOM"; en="M-BOM + MCO"; tag="PLM·API + ERM"; ta=0; d="Mfg BOM + mfg change order"; st="released"; c=$ORANGE},
 @{x=680; y=380; n=9; kr="영향분석·승인"; en="Impact + Approval"; tag="PLM · 승인자"; ta=0; d="Affected objects + VCRM & Checklist 승인"; st="approved"; c=$VIOLET}
)
foreach($k in $cards){ $x=$k.x; $y=$k.y
  Panel $x $y 240 94 $PANEL $k.c 1 0 0.07 | Out-Null
  Bar $x $y 240 4 $k.c 0 | Out-Null
  Oval ($x+8) ($y+8) 18 18 $k.c | Out-Null
  Txt ($x+8) ($y+8) 18 18 ([string]$k.n) 10 -1 $WHITE "Aptos" 2 3 | Out-Null
  Txt ($x+30) ($y+5) 124 15 $k.kr 9.5 -1 $WHITE "Aptos" 1 1 | Out-Null
  if($k.ta -eq 1){ Chip ($x+156) ($y+7) 78 13 $k.tag 6.5 $TAGA_BG $TAGA_TX $TAGA_TX "Aptos" | Out-Null } else { Chip ($x+156) ($y+7) 78 13 $k.tag 6.5 $TAGP_BG $TAGP_TX $TAGP_TX "Aptos" | Out-Null }
  Txt ($x+30) ($y+20) 200 12 $k.en 7 0 $MUTE "Aptos" 1 1 | Out-Null
  Txt ($x+10) ($y+35) 224 26 $k.d 7.5 0 $LIGHT "Aptos" 1 1 | Out-Null
  Chip ($x+10) ($y+68) 150 18 ($k.st + "  OK") 8 "15402F" $GREEN $GREEN "Aptos" | Out-Null
}

Arrow 280 159 360 159 $MUTE 1.75 | Out-Null
Arrow 600 159 680 159 $MUTE 1.75 | Out-Null
Arrow 800 206 800 246 $MUTE 1.75 | Out-Null
Arrow 680 293 600 293 $MUTE 1.75 | Out-Null
Arrow 360 293 280 293 $MUTE 1.75 | Out-Null
Arrow 160 340 160 380 $MUTE 1.75 | Out-Null
Arrow 280 427 360 427 $MUTE 1.75 | Out-Null
Arrow 600 427 680 427 $MUTE 1.75 | Out-Null
$hexPts = @(@(320,159),@(640,159),@(800,226),@(640,293),@(320,293),@(160,360),@(320,427),@(640,427))
foreach($p in $hexPts){ Hex $p[0] $p[1] 46 40 $HEXBG $MAG "AI" $MAG 9.5 | Out-Null }

Txt 36 482 880 14 "AI gate at every step = 오류 체크 + 이상 진단  ·  PASS -> 자동 진행  ·  FAIL -> 차단 & 플래그 (no promote)" 8.5 0 $MAG "Aptos" 1 1 | Out-Null
Txt 36 500 880 14 "권한(RBAC): 설계자 Engineer (작성) · 검토자 Checker (검토) · 승인자 Approver (승인) · 관리자 Admin (프로세스·VCRM 정의)  —  API는 OIDC 역할 스코프로 통제" 7.5 -1 $GOLD "Aptos" 1 1 | Out-Null
Txt 36 518 880 13 "PLM live status:  modeled > DWG issued > MTO > approved > scoped > baselined > ECO closed > released > approved" 7 0 $GREEN "Aptos" 1 1 | Out-Null

$pres.Save()
$flow.Export("$outDir\flow_kr.png","PNG",1280,720)
$pres.Close()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($pres) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($ppt)  | Out-Null
[GC]::Collect(); [GC]::WaitForPendingFinalizers()
Write-Output "BUILD8_FLOW_DONE"
