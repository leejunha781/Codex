$ErrorActionPreference="Stop"

$deck="D:\이력서\AVEVA - Marine Principal Technical Support & Consultant – PLM SME, Busan\Proposal\Future_Industrial_PLM_Meeting_Deck_EN_Rebuilt_v2.pptx"
$reportDir="D:\이력서\AVEVA - Marine Principal Technical Support & Consultant – PLM SME, Busan\Proposal\Validation_Reports"
$stamp=Get-Date -Format "yyyyMMdd_HHmmss"
$report=Join-Path $reportDir ("WBS_KPI_Glossary_Validation_"+$stamp+".md")
$latest=Join-Path $reportDir "WBS_KPI_Glossary_Validation_LATEST.md"
New-Item -ItemType Directory -Force -Path $reportDir|Out-Null

function Norm($s){return (($s-replace"[\r\n\v]"," "-replace"\s+"," ").Trim())}
function Get-SlideText($sl){
 $a=New-Object Collections.Generic.List[string]
 foreach($sh in $sl.Shapes){
  try{
   if($sh.HasTable){
    for($r=1;$r-le$sh.Table.Rows.Count;$r++){for($c=1;$c-le$sh.Table.Columns.Count;$c++){$t=Norm $sh.Table.Cell($r,$c).Shape.TextFrame.TextRange.Text;if($t){$a.Add($t)}}}
   }elseif($sh.HasTextFrame-eq-1-and$sh.TextFrame.HasText-eq-1){$t=Norm $sh.TextFrame.TextRange.Text;if($t){$a.Add($t)}}
  }catch{}
 }
 return $a
}
function Check($cond,$ok,$bad){if($cond){return "PASS — "+$ok}else{return "FAIL — "+$bad}}

$ppt=$null;$pres=$null;$lines=New-Object Collections.Generic.List[string]
try{
 $ppt=New-Object -ComObject PowerPoint.Application;$pres=$ppt.Presentations.Open($deck,$true,$false,$false)
 $wbs=Get-SlideText $pres.Slides.Item(24);$kpi=Get-SlideText $pres.Slides.Item(25)
 $gloss=@();foreach($idx in 31..33){$gloss+=Get-SlideText $pres.Slides.Item($idx)}
 $wbsAll=$wbs-join" | ";$kpiAll=$kpi-join" | ";$glossAll=$gloss-join" | "
 $tables=0;$terms=0
 foreach($idx in 31..33){foreach($sh in $pres.Slides.Item($idx).Shapes){try{if($sh.HasTable){$tables++;$terms+=($sh.Table.Rows.Count-1)}}catch{}}}

 $lines.Add("# WBS · KPI · Glossary Validation Report")
 $lines.Add("")
 $lines.Add("- Generated: "+(Get-Date -Format "yyyy-MM-dd HH:mm:ss K"))
 $lines.Add("- Deck: "+$deck)
 $lines.Add("- Slides: "+$pres.Slides.Count)
 $lines.Add("")
 $lines.Add("## Executive Result")
 $lines.Add("")
 $lines.Add("- WBS: "+(Check ($wbsAll-match"WBS 0"-and$wbsAll-match"WBS 5"-and$wbsAll-match"G1"-and$wbsAll-match"G6") "Six work packages and G1–G6 gates are present." "Required WBS packages or gates are missing."))
 $lines.Add("- KPI: "+(Check ($kpiAll-match"API contract coverage"-and$kpiAll-match"E2E gate pass rate"-and$kpiAll-match"Impact graph recall"-and$kpiAll-match"Dashboard / approval response") "Core configuration and naval-assurance acceptance KPIs are present." "One or more required KPIs are missing."))
 $lines.Add("- Glossary: "+(Check ($tables-eq6-and$terms-ge40) "$tables native tables and $terms glossary entries are present." "Expected 6 native tables and at least 40 entries; found $tables tables and $terms entries."))
 $lines.Add("")
 $lines.Add("## WBS Validation — Slide 24")
 $lines.Add("")
 foreach($req in @("WBS 0","WBS 1","WBS 2","WBS 3","WBS 4","WBS 5","G1","G2","G3","G4","G5","G6","wk26","API contract","E2E test","KPI threshold")){
  $lines.Add("- "+(Check ($wbsAll-match[regex]::Escape($req)) ("Found: "+$req) ("Missing: "+$req)))
 }
 $lines.Add("")
 $lines.Add("## KPI Validation — Slide 25")
 $lines.Add("")
 foreach($req in @("API contract coverage","E2E gate pass rate","BOM / effectivity accuracy","Impact graph recall","Solver-run provenance","Evidence-pack completeness","AI diagnostic precision","Dashboard / approval response","PASS","CONDITIONAL PASS","FAIL")){
  $lines.Add("- "+(Check ($kpiAll-match[regex]::Escape($req)) ("Found: "+$req) ("Missing: "+$req)))
 }
 $lines.Add("")
 $lines.Add("## Glossary Validation — Slides 31–33")
 $lines.Add("")
 $lines.Add("- Native PowerPoint tables: **$tables**")
 $lines.Add("- Glossary data rows: **$terms**")
 foreach($req in @("PLM","BOM","E-BOM","M-BOM","MTO","ECR","ECO","VCRM","NAPA","AVEVA E3D / Marine","API","OpenAPI 3.1","OIDC / OAuth2","RBAC","HA / DR")){
  $lines.Add("- "+(Check ($glossAll-match[regex]::Escape($req)) ("Found: "+$req) ("Missing: "+$req)))
 }
 $lines.Add("")
 $lines.Add("## Recommended Review Actions")
 $lines.Add("")
 $lines.Add("1. Confirm WBS durations and gate ownership before the development-planning meeting.")
 $lines.Add("2. Confirm KPI thresholds with development, planning, naval architecture and class stakeholders.")
 $lines.Add("3. Add new abbreviations to the native Glossary tables whenever new API, solver or infrastructure terms are introduced.")
 $lines.Add("4. Re-run this report after material changes to slides 24, 25 or 31–33.")
}finally{if($pres){$pres.Close()};if($ppt){if($ppt.Presentations.Count-eq0){$ppt.Quit()}}}

$body=$lines-join"`r`n"
[IO.File]::WriteAllText($report,$body,[Text.Encoding]::UTF8)
[IO.File]::WriteAllText($latest,$body,[Text.Encoding]::UTF8)
Write-Output("REPORT="+$report)
Write-Output("LATEST="+$latest)
