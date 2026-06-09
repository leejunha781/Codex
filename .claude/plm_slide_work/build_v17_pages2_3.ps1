$ErrorActionPreference = "Stop"

$dir   = "D:\이력서\AVEVA - Marine Principal Technical Support & Consultant – PLM SME, Busan\Proposal\"
$v16   = $dir + "Future_Industrial_PLM_Meeting_Deck_EN_V16.pptx"
$v17   = $dir + "Future_Industrial_PLM_Meeting_Deck_EN_V17.pptx"
$alias = $dir + "Future_Industrial_PLM_Meeting_Deck_EN.pptx"
$final = $dir + "Future_Industrial_PLM_Meeting_Deck_EN_Final.pptx"
$fpdf  = $dir + "Future_Industrial_PLM_Meeting_Deck_EN_Final.pdf"
$bk    = $dir + "_backup_20260609_v17_pages2_3_visual\"

New-Item -ItemType Directory -Force -Path $bk | Out-Null
foreach ($f in @($v16,$alias,$final)) { if (Test-Path $f) { Copy-Item $f -Destination $bk -Force } }
Copy-Item $v16 -Destination $v17 -Force

# chars + palette
$nd=[char]0x2013; $md=[char]0x00B7; $ar=[char]0x2192; $ge=[char]0x2265; $le=[char]0x2264; $CR=[char]13
$navy=3678739; $panel=4205084; $cyan=15646772; $amber=3846888; $green=9224777
$white=16777215; $ltgray=15062727; $muted=12626067; $violet=16419751
$blue=16155195; $magenta=11161814; $teal=12571693; $orange=1471481; $rose=11956980
$titleFont="Aptos Display"; $bodyFont="Aptos"

$pp=$null
for($a=0;$a -lt 6 -and $pp -eq $null;$a++){ try{$pp=New-Object -ComObject PowerPoint.Application}catch{Start-Sleep -Milliseconds 1200} }
if($pp -eq $null){ throw "PowerPoint COM did not start" }
try { $pp.Visible = $true } catch {}

$pres = $pp.Presentations.Open($v17, $false, $false, $false)
$W = $pres.PageSetup.SlideWidth
$H = $pres.PageSetup.SlideHeight

function Add-Box($sl,$l,$t,$w,$h,$text,$size,$bold,$color,$font,$align){
    $tb=$sl.Shapes.AddTextbox(1,[single]$l,[single]$t,[single]$w,[single]$h)
    $tb.TextFrame.WordWrap=-1; $tb.TextFrame.AutoSize=0
    $tb.TextFrame.MarginLeft=[single]2; $tb.TextFrame.MarginRight=[single]2
    $tb.TextFrame.MarginTop=[single]2; $tb.TextFrame.MarginBottom=[single]2
    $tr=$tb.TextFrame.TextRange
    $tr.Text=$text; $tr.Font.Size=[single]$size; $tr.Font.Bold=[int]$bold
    $tr.Font.Name=$font; $tr.Font.Color.RGB=$color; $tr.ParagraphFormat.Alignment=[int]$align
    return $tb
}
function Add-Panel($sl,$l,$t,$w,$h,$fill,$line,$lw){
    $p=$sl.Shapes.AddShape(5,[single]$l,[single]$t,[single]$w,[single]$h)
    $p.Fill.Solid(); $p.Fill.ForeColor.RGB=$fill
    $p.Line.Visible=-1; $p.Line.ForeColor.RGB=$line; $p.Line.Weight=[single]$lw
    $p.Shadow.Visible=0
    return $p
}
function Add-Line($sl,$l,$t,$w,$color){
    $r=$sl.Shapes.AddShape(1,[single]$l,[single]$t,[single]$w,[single]1.4)
    $r.Fill.Solid(); $r.Fill.ForeColor.RGB=$color; $r.Fill.Transparency=[single]0.45
    $r.Line.Visible=0; $r.Shadow.Visible=0
    return $r
}
function Prep-Slide($sl){
    foreach($sh in @($sl.Shapes)){ $sh.Delete() }
    $sl.FollowMasterBackground=$false
    $sl.Background.Fill.Solid(); $sl.Background.Fill.ForeColor.RGB=$navy
    return $sl
}

# ============== SLIDE 2 : RUN-OF-SHOW (improved) ==============
$s = Prep-Slide ($pres.Slides.Item(2))
Add-Box $s 40 22 760 36 "Technical meeting agenda" 26 1 $white $titleFont 1 | Out-Null
$subR="Future Industrial PLM " + $md + " AVEVA Marine " + $md + " Development + Planning decision meeting " + $md + " 90 minutes"
Add-Box $s 40 60 ($W-80) 22 $subR 11.5 0 $ltgray $bodyFont 1 | Out-Null
Add-Box $s ($W-330) 18 250 18 "Future Industrial PLM" 9 0 $muted $bodyFont 3 | Out-Null
Add-Box $s ($W-52) 16 38 18 "2" 11 0 $muted $bodyFont 3 | Out-Null

$ban=Add-Panel $s 40 92 ($W-80) 48 $panel $cyan 1.5
$banTxt="DECISION REQUESTED TODAY  " + $md + "  Approve Phase 2 configuration core + one bounded Continuous Naval Assurance pilot " + $ar + " named owners, KPI thresholds, assumption boundaries and the first evidence scenario."
$bt=Add-Box $s 54 98 ($W-108) 38 $banTxt 11 1 $white $bodyFont 1
$bt.TextFrame.VerticalAnchor=3

$rows=@(
 @("#","Min","Agenda block","Deck","Lead " + $ar + " audience","Outcome to capture / decide"),
 @("0","5","Open & decision objective","p.1",("Facilitator " + $ar + " all"),"Confirm the exact approval requested today"),
 @("1","5",("How we'll decide " + $nd + " procedure"),("p.4" + $nd + "6"),("Facilitator " + $md + " PLM SME " + $ar + " all"),"Agree decision flow; set dev vs planning lenses"),
 @("2","10","Current state & why now",("p.7" + $nd + "10"),("PLM SME " + $ar + " both teams"),"Accept that the gap and urgency are real"),
 @("3","15","Winning ideas / strategy",("p.11" + $nd + "19"),("PLM SME " + $ar + " both teams"),"Agree win-above position + the AI stance"),
 @("4","20","Proposal & architecture",("p.20" + $nd + "30"),("PLM SME + Architect " + $ar + " Dev"),"Confirm hybrid architecture is feasible & bounded"),
 @("5","10",("Delivery " + $nd + " WBS & KPI gates"),("p.31" + $nd + "33"),("PLM SME " + $ar + " Planning"),"Agree the 26-wk staged plan + KPI acceptance"),
 @("6","5","Future-state payoff","p.34",("PLM SME " + $ar + " Planning"),"Judge the competitive upside credible"),
 @("7","10",("Review " + $nd + " ownership, risks, evidence"),("p.35" + $nd + "37"),("Facilitator " + $ar + " all"),"Assign owners; accept risks & assumptions"),
 @("8","10","Decision & close",("p.38" + $nd + "39"),("Facilitator " + $ar + " all"),"Record decision, scope, owners, first scenario")
)
$acc=@($cyan,$green,$blue,$magenta,$amber,$teal,$orange,$violet,$rose)

$tLeft=40; $tTop=150; $tWidth=($W-80); $nRows=$rows.Count; $nCols=6
$tblShape=$s.Shapes.AddTable($nRows,$nCols,[single]$tLeft,[single]$tTop,[single]$tWidth,[single]340)
$tbl=$tblShape.Table
$frac=@(0.038,0.050,0.235,0.075,0.215,0.387)
for($c=1;$c -le $nCols;$c++){ $tbl.Columns.Item($c).Width=[single]($tWidth*$frac[$c-1]) }
$aligns=@(2,2,1,2,1,1)
for($r=1;$r -le $nRows;$r++){
    $tbl.Rows.Item($r).Height=[single]34
    for($c=1;$c -le $nCols;$c++){
        $cell=$tbl.Cell($r,$c)
        $cell.Shape.TextFrame.VerticalAnchor=3
        $cell.Shape.TextFrame.MarginLeft=[single]6; $cell.Shape.TextFrame.MarginRight=[single]6
        $cell.Shape.TextFrame.MarginTop=[single]2; $cell.Shape.TextFrame.MarginBottom=[single]2
        $tr=$cell.Shape.TextFrame.TextRange
        $tr.Text=[string]$rows[$r-1][$c-1]; $tr.Font.Name=$bodyFont
        $tr.ParagraphFormat.Alignment=[int]$aligns[$c-1]
        for($b=1;$b -le 4;$b++){ try{ $cell.Borders.Item($b).ForeColor.RGB=$navy; $cell.Borders.Item($b).Weight=[single]1 }catch{} }
        if($r -eq 1){
            $cell.Shape.Fill.ForeColor.RGB=$panel
            $tr.Font.Size=[single]11; $tr.Font.Bold=[int]1; $tr.Font.Color.RGB=$cyan
        } else {
            if($r % 2 -eq 0){ $cell.Shape.Fill.ForeColor.RGB=$navy } else { $cell.Shape.Fill.ForeColor.RGB=$panel }
            $tr.Font.Size=[single]9.5; $tr.Font.Bold=[int]0; $tr.Font.Color.RGB=$ltgray
            if($c -eq 1){ $cell.Shape.Fill.ForeColor.RGB=$acc[$r-2]; $tr.Font.Bold=[int]1; $tr.Font.Color.RGB=$navy; $tr.Font.Size=[single]11 }
            if($c -eq 2){ $tr.Font.Color.RGB=$white; $tr.Font.Bold=[int]1 }
            if($c -eq 3){ $tr.Font.Bold=[int]1; $tr.Font.Color.RGB=$white; $tr.Font.Size=[single]10.5 }
            if($c -eq 4){ $tr.Font.Color.RGB=$cyan; $tr.Font.Bold=[int]1 }
            if($c -eq 5){ $tr.Font.Color.RGB=$white; $tr.Font.Size=[single]9.5 }
            if($c -eq 6){ $tr.Font.Color.RGB=$ltgray; $tr.Font.Size=[single]9 }
        }
    }
}

$legend="PLM SME = Principal Technical Support & Consultant (presenter)   " + $md + "   Facilitator = chair / decision owner   " + $md + "   Architect = Solution Architect   " + $md + "   Glossary p.40" + $nd + "42 = on-demand reference"
Add-Box $s 40 497 ($W-80) 16 $legend 8.5 0 $muted $bodyFont 1 | Out-Null

# ============== SLIDE 3 : HOW-TO-RUN (improved readability) ==============
$h2=Prep-Slide ($pres.Slides.Item(3))
$howTitle="How to run the room " + $nd + " roles, decision criteria & ownership"
Add-Box $h2 40 22 ($W-80) 30 $howTitle 22 1 $white $titleFont 1 | Out-Null
Add-Box $h2 ($W-330) 18 250 18 "Future Industrial PLM" 9 0 $muted $bodyFont 3 | Out-Null
Add-Box $h2 ($W-52) 16 38 18 "3" 11 0 $muted $bodyFont 3 | Out-Null

$pTop1=64; $pH1=222; $pTop2=296; $pH2=224
$cLW=40; $cLWd=430; $cRW=490; $cRWd=430

# Panel A : roles & pre-reads
$null=Add-Panel $h2 $cLW $pTop1 $cLWd $pH1 $panel $cyan 1.4
Add-Box $h2 ($cLW+16) ($pTop1+10) ($cLWd-32) 18 "ROLES & PRE-READS" 11.5 1 $cyan $bodyFont 1 | Out-Null
$rolesTxt="Chair / decision owner " + $md + " runs the room, holds the decision" + $CR +
          "PLM SME (presenter) " + $md + " walks the narrative" + $CR +
          "Solution Architect " + $md + " architecture & feasibility" + $CR +
          "Development lead " + $md + " feasibility, sequencing, estimates" + $CR +
          "Planning lead " + $md + " position, value, scope, gates" + $CR +
          "Naval architecture / Class " + $md + " solver & evidence acceptance" + $CR +
          "Scribe " + $md + " logs decisions, owners, actions"
Add-Box $h2 ($cLW+16) ($pTop1+36) ($cLWd-32) 122 $rolesTxt 10.5 0 $ltgray $bodyFont 1 | Out-Null
Add-Line $h2 ($cLW+16) ($pTop1+162) ($cLWd-32) $cyan | Out-Null
$preTxt="Pre-reads (24" + $nd + "48h before): p.1" + $nd + "2, p.20" + $nd + "23, p.32" + $nd + "33, p.38; Glossary p.40" + $nd + "42." + $CR +
        "Ground rule (p.4): not a tool-replacement story " + $md + " a governed, auditable, operations-aware decision loop."
Add-Box $h2 ($cLW+16) ($pTop1+170) ($cLWd-32) 50 $preTxt 9.5 0 $muted $bodyFont 1 | Out-Null

# Panel B : decision criteria
$null=Add-Panel $h2 $cRW $pTop1 $cRWd $pH1 $panel $green 1.4
Add-Box $h2 ($cRW+16) ($pTop1+10) ($cRWd-32) 18 "DECISION CRITERIA" 11.5 1 $green $bodyFont 1 | Out-Null
$kpi="contract 100% " + $md + " config " + $ge + "99% " + $md + " impact recall " + $ge + "95% " + $md + " provenance 100% " + $md + " approval " + $le + "5s"
$apTxt="APPROVE " + $ar + " Phase 2 configuration core + one bounded Continuous Naval Assurance pilot, with named owners; KPI thresholds (" + $kpi + "); assumption boundaries; and one pilot scenario (one hull/block " + $md + " selected loading cases " + $md + " one approved solver adapter " + $md + " seeded change & failure cases)."
Add-Box $h2 ($cRW+16) ($pTop1+36) ($cRWd-32) 120 $apTxt 10 0 $ltgray $bodyFont 1 | Out-Null
Add-Line $h2 ($cRW+16) ($pTop1+162) ($cRWd-32) $green | Out-Null
$ccTxt="CONDITIONAL " + $ar + " approve with a named backlog to clear before G1/G2." + $CR +
       "DEFER " + $ar + " list the specific blockers + the evidence needed to revisit."
Add-Box $h2 ($cRW+16) ($pTop1+170) ($cRWd-32) 50 $ccTxt 10 0 $white $bodyFont 1 | Out-Null

# Panel C : development owns
$null=Add-Panel $h2 $cLW $pTop2 $cLWd $pH2 $panel $amber 1.4
Add-Box $h2 ($cLW+16) ($pTop2+10) ($cLWd-32) 18 "DEVELOPMENT OWNS" 11.5 1 $amber $bodyFont 1 | Out-Null
$buildTxt="BUILD: OpenAPI contracts (objects, BOM, effectivity, baseline, change impact, evidence); Local Agent / Plugin callbacks from Windows authoring; AI-gate provider interface (rules " + $ar + " RAG " + $ar + " CONNECT); seed truth cases; E2E tests; KPI gates."
Add-Box $h2 ($cLW+16) ($pTop2+36) ($cLWd-32) 110 $buildTxt 10.5 0 $ltgray $bodyFont 1 | Out-Null
Add-Line $h2 ($cLW+16) ($pTop2+150) ($cLWd-32) $amber | Out-Null
$avoidTxt="AVOID: E3D / Marine / Hull on Linux; big-bang solver replacement; automatic AI release authority; unbounded process customization without approval gates."
Add-Box $h2 ($cLW+16) ($pTop2+158) ($cLWd-32) 60 $avoidTxt 10.5 0 $white $bodyFont 1 | Out-Null

# Panel D : planning owns + capture
$null=Add-Panel $h2 $cRW $pTop2 $cRWd $pH2 $panel $violet 1.4
Add-Box $h2 ($cRW+16) ($pTop2+10) ($cRWd-32) 18 "PLANNING OWNS  +  CAPTURE LIVE" 11.5 1 $violet $bodyFont 1 | Out-Null
$planTxt="PLANNING OWNS: product position (control plane above tools); the customer-value story; 26-week gate discipline; the decision ask; which process variants are allowed."
Add-Box $h2 ($cRW+16) ($pTop2+36) ($cRWd-32) 96 $planTxt 10.5 0 $ltgray $bodyFont 1 | Out-Null
Add-Line $h2 ($cRW+16) ($pTop2+150) ($cRWd-32) $violet | Out-Null
$capTxt="CAPTURE LIVE: Decision log " + $md + " Action log (action " + $md + " owner " + $md + " due " + $md + " gate) " + $md + " Parking lot for off-scope items (record, don't debate)."
Add-Box $h2 ($cRW+16) ($pTop2+158) ($cRWd-32) 60 $capTxt 10.5 0 $white $bodyFont 1 | Out-Null

# save
$pres.Save()
$pres.SaveAs($alias)
$pres.SaveAs($final)
$pres.SaveAs($fpdf, 32)
$pres.Close()
$pp.Quit()
Write-Output "V17 saved (pages 2-3 reworked) + alias/final/pdf"
