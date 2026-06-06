$ErrorActionPreference = "Stop"
$dir = "D:\이력서\AVEVA - Marine Principal Technical Support & Consultant – PLM SME, Busan\Proposal"
$v2 = Join-Path $dir "Future_Industrial_PLM_Meeting_Deck_EN_Rebuilt_v2.pptx"
$out = Join-Path $dir "AVEVA_비교분석_Comparison.pptx"
$outName = "AVEVA_비교분석_Comparison.pptx"
$png = "C:\Users\namma\.claude\plm_slide_work\review_cmp"
New-Item -ItemType Directory -Force -Path $png | Out-Null

Copy-Item -LiteralPath $v2 -Destination $out -Force

function Rgb([string]$hex){ $hex=$hex.TrimStart('#'); [int]([Convert]::ToInt32($hex.Substring(0,2),16)+[Convert]::ToInt32($hex.Substring(2,2),16)*256+[Convert]::ToInt32($hex.Substring(4,2),16)*65536) }
function Retry($sb){ for($a=0;$a -lt 6;$a++){ try{ return (& $sb) }catch{ Start-Sleep -Milliseconds 80 } }; return (& $sb) }
$WHITE="FFFFFF"; $LIGHT="C7D6E5"; $MUTE="93A8C0"; $CYAN="34C0EE"; $GREEN="49C28C"; $BGD="0B1626"

$ppt = New-Object -ComObject PowerPoint.Application
foreach($p in @($ppt.Presentations)){ if($p.Name -eq $outName){ try{ $p.Saved=$true; $p.Close() }catch{} } }
$pres = $ppt.Presentations.Open($out,0,0,0)

$keep = @("Current PLM landscape","Public roadmap signals","Strategic gap to fill","Benchmark absorption map","Evidence base")
function HasTitle($sl,$titles){ foreach($sh in $sl.Shapes){ try{ if($sh.HasTextFrame -eq -1 -and $sh.TextFrame.HasText -eq -1){ $t=$sh.TextFrame.TextRange.Text.Trim(); foreach($k in $titles){ if($t -eq $k){ return $true } } } }catch{} }; return $false }

$kept=@()
for($i=$pres.Slides.Count; $i -ge 1; $i--){
  $sl=$pres.Slides.Item($i)
  if(HasTitle $sl $keep){ $kept += $i } else { $sl.Delete() }
}
Write-Output ("kept slide count = " + $pres.Slides.Count)

# cover slide at front
$script:slide = $pres.Slides.Add(1,12)
function NoShadow($s){ try{$s.Shadow.Visible=0}catch{} }
function Txt($l,$t,$w,$h,$text,$size,$bold,$color,$font,$align,$anchor){ $tb=$script:slide.Shapes.AddTextbox(1,[single]$l,[single]$t,[single]$w,[single]$h); $tb.Fill.Visible=0;$tb.Line.Visible=0;NoShadow $tb; $tf=$tb.TextFrame;$tf.WordWrap=-1;$tf.AutoSize=0;$tf.VerticalAnchor=[int]$anchor; $tr=$tf.TextRange;$tr.Text=$text;$tr.Font.Size=[single]$size;$tr.Font.Bold=[int]$bold;$tr.Font.Name=[string]$font;Retry{$tr.Font.Color.RGB=[int](Rgb $color)};$tr.ParagraphFormat.Alignment=[int]$align; return $tb }
$bg=$script:slide.Shapes.AddShape(1,[single]0,[single]0,[single]960,[single]540); $bg.Line.Visible=0; NoShadow $bg; $bg.Fill.Solid(); Retry{$bg.Fill.ForeColor.RGB=[int](Rgb $BGD)}
$rb=$script:slide.Shapes.AddShape(1,[single]60,[single]250,[single]120,[single]5); $rb.Line.Visible=0; $rb.Fill.Solid(); Retry{$rb.Fill.ForeColor.RGB=[int](Rgb $CYAN)}; NoShadow $rb
Txt 60 150 840 22 "BENCHMARK / 비교분석" 13 -1 $CYAN "Aptos" 1 1 | Out-Null
Txt 58 180 850 60 "AVEVA vs 타 PLM · 조선 설계툴 — 비교분석" 30 -1 $WHITE "Aptos Display" 1 1 | Out-Null
Txt 60 270 850 22 "현재 상태 + 로드맵(Roadmap) 기준  ·  HD중공업 수주 검토용" 14 0 $LIGHT "Aptos" 1 1 | Out-Null
Txt 60 320 850 20 "① 현재 PLM 지형  ②공개 로드맵 시그널  ③전략적 공백  ④벤치마크 흡수맵  ⑤근거 자료" 11 0 $MUTE "Aptos" 1 1 | Out-Null
Txt 60 470 850 16 "※ 공개 출처 기반 내부 검토용 · Siemens / Dassault / Hexagon / PTC vs AVEVA" 9 0 $MUTE "Aptos" 1 1 | Out-Null

$pres.SaveAs($out, 24)
for($i=1;$i -le $pres.Slides.Count;$i++){ $num="{0:D2}" -f $i; $pres.Slides.Item($i).Export("$png\cmp-$num.png","PNG",1280,720) }
Write-Output ("FINAL compare slides = " + $pres.Slides.Count)
$pres.Close()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($pres) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($ppt) | Out-Null
[GC]::Collect(); [GC]::WaitForPendingFinalizers()
Write-Output ("COMPARE_DONE " + $out)
