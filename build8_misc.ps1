$ErrorActionPreference = "Stop"
$v2 = "D:\이력서\AVEVA - Marine Principal Technical Support & Consultant – PLM SME, Busan\Proposal\Future_Industrial_PLM_Meeting_Deck_EN_Rebuilt_v2.pptx"
$name = "Future_Industrial_PLM_Meeting_Deck_EN_Rebuilt_v2.pptx"
$outDir = "C:\Users\namma\.claude\plm_slide_work\review_png8"

function Rgb([string]$hex){ $hex=$hex.TrimStart('#'); [int]([Convert]::ToInt32($hex.Substring(0,2),16)+[Convert]::ToInt32($hex.Substring(2,2),16)*256+[Convert]::ToInt32($hex.Substring(4,2),16)*65536) }
function Retry($sb){ for($a=0;$a -lt 6;$a++){ try{ return (& $sb) }catch{ Start-Sleep -Milliseconds 80 } }; return (& $sb) }
$WHITE="FFFFFF"; $CYAN="34C0EE"; $ORANGE="F0832F"; $VIOLET="9B7BE0"; $MUTE="93A8C0"

$ppt = New-Object -ComObject PowerPoint.Application
foreach($p in @($ppt.Presentations)){ if($p.Name -eq $name){ try{ $p.Saved=$true; $p.Close() }catch{} } }
$pres=$ppt.Presentations.Open($v2,0,0,0)

function EditByPrefix($prefix,$new,$size,$bold,$color,$font){
  foreach($sl in $pres.Slides){ foreach($sh in @($sl.Shapes)){ try{ if($sh.HasTextFrame -eq -1 -and $sh.TextFrame.HasText -eq -1){ if($sh.TextFrame.TextRange.Text.Trim().StartsWith($prefix)){ $r=$sh.TextFrame.TextRange; $r.Text=$new; $r.Font.Size=[single]$size; $r.Font.Bold=[int]$bold; $r.Font.Name=[string]$font; Retry{$r.Font.Color.RGB=[int](Rgb $color)}; return $sl.SlideIndex } } }catch{} } }; return 0
}

# --- Slide 12 (VCRM): bilingual title, headers, gate ---
$a=EditByPrefix "Requirement ->" "요구사항 Requirement -> AI VCRM·체크리스트  (관리자 편집 · Impact+Approval 승인)" 14 -1 $WHITE "Aptos Display"
$b=EditByPrefix "VCRM —" "VCRM · 검증요구사항 매트릭스 (대분류)" 9 -1 $CYAN "Aptos"
$c=EditByPrefix "Checklist —" "체크리스트 Checklist (대분류) — 유의사항 · 반영확인" 9 -1 $ORANGE "Aptos"
$d=EditByPrefix "APPROVED IN" "Impact + Approval 단계에서 일괄 승인" 8.5 -1 $VIOLET "Aptos"
Write-Output ("S12: title@$a vcrm@$b chk@$c gate@$d")

# --- Slide 14 (drawing): bilingual title ---
$e=EditByPrefix "Design tool ->" "설계툴 -> 2D 생산도면 (AutoCAD DWG/DXF) -> API PLM 등록" 16 -1 $WHITE "Aptos Display"
Write-Output ("S14: title@$e")

# --- Slide 15 (process): bilingual title + admin->관리자 ---
$f=EditByPrefix "User-configurable" "사용자 구성형 PLM 프로세스 — API로 전체 흐름 변경 (관리자 권한)" 16 -1 $WHITE "Aptos Display"
$g=EditByPrefix "Future PLM  —  Process Designer" "Future PLM  —  Process Designer  ·  관리자 admin  ·  /api/v1" 8 0 $MUTE "Consolas"
Write-Output ("S15: title@$f bar@$g")

$pres.Save()
foreach($pp in @($a,$e,$f)){ if($pp -gt 0){ $pres.Slides.Item($pp).Export("$outDir\s$pp.png","PNG",1280,720) } }
$pres.Close()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($pres) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($ppt)  | Out-Null
[GC]::Collect(); [GC]::WaitForPendingFinalizers()
Write-Output "BUILD8_MISC_DONE"
