$ErrorActionPreference = "Stop"

$dir   = "D:\이력서\AVEVA - Marine Principal Technical Support & Consultant – PLM SME, Busan\Proposal\"
$v15   = $dir + "Future_Industrial_PLM_Meeting_Deck_EN_V15.pptx"
$v16   = $dir + "Future_Industrial_PLM_Meeting_Deck_EN_V16.pptx"
$alias = $dir + "Future_Industrial_PLM_Meeting_Deck_EN.pptx"
$final = $dir + "Future_Industrial_PLM_Meeting_Deck_EN_Final.pptx"
$fpdf  = $dir + "Future_Industrial_PLM_Meeting_Deck_EN_Final.pdf"
$bk    = $dir + "_backup_20260609_v16_agenda_integration\"

# ---- backup ----
New-Item -ItemType Directory -Force -Path $bk | Out-Null
foreach ($f in @($v15,$alias,$final)) { if (Test-Path $f) { Copy-Item $f -Destination $bk -Force } }
# ---- create V16 working copy from V15 ----
Copy-Item $v15 -Destination $v16 -Force

# ---- special chars + palette ----
$nd=[char]0x2013; $md=[char]0x00B7; $ar=[char]0x2192; $ge=[char]0x2265; $le=[char]0x2264; $CR=[char]13
$navy=3678739; $panel=4205084; $cyan=15646772; $amber=3846888; $green=9224777
$white=16777215; $ltgray=15062727; $muted=12626067; $violet=16419751
$titleFont="Aptos Display"; $bodyFont="Aptos"

# ---- page-number shape map (old slide index -> shape name) ----
$map = @{ 3="TextBox 44"; 4="TextBox 40"; 5="TextBox 44"; 6="TextBox 59"; 7="Text 10"; 8="Text 10";
 9="TextBox 50"; 10="Text 7"; 11="TextBox 79"; 12="TextBox 119"; 13="TextBox 5"; 14="TextBox 63";
 15="TextBox 63"; 16="TextBox 26"; 17="TextBox 26"; 18="TextBox 26"; 19="Text 9"; 20="Text 12";
 21="Text 12"; 22="Text 12"; 23="TextBox 89"; 24="Text 13"; 25="Text 13"; 26="Text 13"; 27="Text 13";
 28="TextBox 5"; 29="Text 13"; 30="Text 9"; 31="TextBox 119"; 32="TextBox 125"; 33="TextBox 117";
 34="TextBox 57"; 35="TextBox 116"; 36="TextBox 80"; 37="TextBox 65"; 38="Slide38PageNumV11";
 39="TextBox 10"; 40="TextBox 10"; 41="TextBox 10" }

# ---- COM ----
$pp = $null
for ($a=0; $a -lt 6 -and $pp -eq $null; $a++) { try { $pp = New-Object -ComObject PowerPoint.Application } catch { Start-Sleep -Milliseconds 1200 } }
if ($pp -eq $null) { throw "PowerPoint COM did not start" }
try { $pp.Visible = $true } catch {}

$pres = $pp.Presentations.Open($v16, $false, $false, $false)
$W = $pres.PageSetup.SlideWidth
$H = $pres.PageSetup.SlideHeight

# ---- helpers ----
function Add-Box($sl,$l,$t,$w,$h,$text,$size,$bold,$color,$font,$align) {
    $tb = $sl.Shapes.AddTextbox(1, [single]$l, [single]$t, [single]$w, [single]$h)
    $tb.TextFrame.WordWrap = -1; $tb.TextFrame.AutoSize = 0
    $tb.TextFrame.MarginLeft=[single]2; $tb.TextFrame.MarginRight=[single]2
    $tb.TextFrame.MarginTop=[single]2; $tb.TextFrame.MarginBottom=[single]2
    $tr = $tb.TextFrame.TextRange
    $tr.Text=$text; $tr.Font.Size=[single]$size; $tr.Font.Bold=[int]$bold
    $tr.Font.Name=$font; $tr.Font.Color.RGB=$color; $tr.ParagraphFormat.Alignment=[int]$align
    return $tb
}
function Add-Panel($sl,$l,$t,$w,$h,$fill,$line,$lw) {
    $p = $sl.Shapes.AddShape(5,[single]$l,[single]$t,[single]$w,[single]$h)
    $p.Fill.Solid(); $p.Fill.ForeColor.RGB=$fill
    $p.Line.Visible=-1; $p.Line.ForeColor.RGB=$line; $p.Line.Weight=[single]$lw
    $p.Shadow.Visible=0
    return $p
}
function Prep-Slide($sl) {
    foreach ($sh in @($sl.Shapes)) { $sh.Delete() }
    $sl.FollowMasterBackground = $false
    $sl.Background.Fill.Solid(); $sl.Background.Fill.ForeColor.RGB = $navy
    return $sl
}

# ===== STEP 1: renumber old slides 3..41 (+1), in place =====
$mismatch = @()
for ($i=3; $i -le 41; $i++) {
    $sl = $pres.Slides.Item($i)
    $nm = $map[$i]
    try {
        $shp = $sl.Shapes.Item($nm)
        $cur = ($shp.TextFrame.TextRange.Text -replace "\s","")
        if ($cur -eq [string]$i) { $shp.TextFrame.TextRange.Text = [string]($i+1) }
        else { $mismatch += ("slide $i shape '$nm' text='$cur' (expected $i)") }
    } catch { $mismatch += ("slide $i shape '$nm' NOT FOUND") }
}

# ===== STEP 2: fix slide 18 WBS reference =====
try {
    $s18 = $pres.Slides.Item(18).Shapes.Item("TextBox 177")
    $t18 = $s18.TextFrame.TextRange.Text
    $s18.TextFrame.TextRange.Text = ($t18 -replace "WBS p\.31","WBS p.32")
} catch { $mismatch += "slide18 TextBox 177 ref fix failed" }

# ===== STEP 3: cover revision -> V16 =====
try { $pres.Slides.Item(1).Shapes.Item("RevisionStamp").TextFrame.TextRange.Text = "Rev. V16 " + $md + " 2026-06-09" } catch { $mismatch += "RevisionStamp fail" }

# ===== STEP 4: delete old section-map (slide 2) =====
$pres.Slides.Item(2).Delete()

# ===== STEP 5: insert two blank slides at positions 2 and 3 =====
$lay = $pres.Slides.Item(1).CustomLayout
$null = $pres.Slides.AddSlide(2, $lay)
$null = $pres.Slides.AddSlide(3, $lay)
$sRun = Prep-Slide ($pres.Slides.Item(2))
$sHow = Prep-Slide ($pres.Slides.Item(3))

# ===== STEP 6: build RUN-OF-SHOW on slide 2 =====
Add-Box $sRun 40 22 760 36 "Technical meeting agenda" 26 1 $white $titleFont 1 | Out-Null
$subR = "Future Industrial PLM " + $md + " AVEVA Marine " + $md + " Development + Planning decision meeting " + $md + " 90 minutes"
Add-Box $sRun 40 60 ($W-80) 22 $subR 11.5 0 $ltgray $bodyFont 1 | Out-Null
Add-Box $sRun ($W-330) 18 250 18 "Future Industrial PLM" 9 0 $muted $bodyFont 3 | Out-Null
Add-Box $sRun ($W-52) 16 38 18 "2" 11 0 $muted $bodyFont 3 | Out-Null

$ban = Add-Panel $sRun 40 92 ($W-80) 50 $panel $cyan 1.5
$banTxt = "DECISION REQUESTED TODAY  " + $md + "  Approve Phase 2 configuration core + one bounded Continuous Naval Assurance pilot " + $ar + " named owners, KPI thresholds, assumption boundaries and the first evidence scenario."
$bt = Add-Box $sRun 54 99 ($W-108) 38 $banTxt 11 1 $white $bodyFont 1
$bt.TextFrame.VerticalAnchor = 3

$rows = @(
 @("#","Min","Agenda block","Deck","Lead","Outcome to capture / decide"),
 @("0","5","Open & decision objective","p.1","Chair","Confirm the exact approval requested today"),
 @("1","5",("How we'll decide " + $nd + " procedure"),("p.4" + $nd + "6"),("Chair " + $md + " SME"),"Agree decision flow; set dev vs planning lenses"),
 @("2","10","Current state & why now",("p.7" + $nd + "10"),"SME","Accept that the gap and urgency are real"),
 @("3","15","Winning ideas / strategy",("p.11" + $nd + "19"),"SME","Agree win-above position + the AI stance"),
 @("4","20","Proposal & architecture",("p.20" + $nd + "30"),("SME " + $md + " Architect"),"Confirm hybrid architecture is feasible & bounded"),
 @("5","10",("Delivery " + $nd + " WBS & KPI gates"),("p.31" + $nd + "33"),"SME","Agree the 26-wk staged plan + KPI acceptance"),
 @("6","5","Future-state payoff","p.34","SME","Judge the competitive upside credible"),
 @("7","10",("Review " + $nd + " ownership, risks, evidence"),("p.35" + $nd + "37"),"Chair","Assign owners; accept risks & assumptions"),
 @("8","10","Decision & close",("p.38" + $nd + "39"),"Chair","Record decision, scope, owners, first scenario")
)
$tLeft=40; $tTop=152; $tWidth=($W-80); $nRows=$rows.Count; $nCols=6
$tblShape = $sRun.Shapes.AddTable($nRows,$nCols,[single]$tLeft,[single]$tTop,[single]$tWidth,[single]340)
$tbl = $tblShape.Table
$frac=@(0.040,0.052,0.262,0.082,0.150,0.414)
for ($c=1;$c -le $nCols;$c++){ $tbl.Columns.Item($c).Width=[single]($tWidth*$frac[$c-1]) }
$aligns=@(2,2,1,2,1,1)
for ($r=1;$r -le $nRows;$r++){
    $tbl.Rows.Item($r).Height=[single]34
    for ($c=1;$c -le $nCols;$c++){
        $cell=$tbl.Cell($r,$c)
        $cell.Shape.TextFrame.VerticalAnchor=3
        $cell.Shape.TextFrame.MarginLeft=[single]6; $cell.Shape.TextFrame.MarginRight=[single]6
        $cell.Shape.TextFrame.MarginTop=[single]2; $cell.Shape.TextFrame.MarginBottom=[single]2
        $tr=$cell.Shape.TextFrame.TextRange
        $tr.Text=[string]$rows[$r-1][$c-1]; $tr.Font.Name=$bodyFont
        $tr.ParagraphFormat.Alignment=[int]$aligns[$c-1]
        for ($b=1;$b -le 4;$b++){ try { $cell.Borders.Item($b).ForeColor.RGB=$navy; $cell.Borders.Item($b).Weight=[single]1 } catch {} }
        if ($r -eq 1){ $cell.Shape.Fill.ForeColor.RGB=$panel; $tr.Font.Size=[single]10.5; $tr.Font.Bold=[int]1; $tr.Font.Color.RGB=$cyan }
        else {
            if ($r % 2 -eq 0){ $cell.Shape.Fill.ForeColor.RGB=$navy } else { $cell.Shape.Fill.ForeColor.RGB=$panel }
            $tr.Font.Size=[single]9; $tr.Font.Bold=[int]0; $tr.Font.Color.RGB=$ltgray
            if ($c -eq 1){ $tr.Font.Bold=[int]1; $tr.Font.Color.RGB=$cyan; $tr.Font.Size=[single]10 }
            if ($c -eq 2){ $tr.Font.Color.RGB=$white }
            if ($c -eq 3){ $tr.Font.Bold=[int]1; $tr.Font.Color.RGB=$white; $tr.Font.Size=[single]9.5 }
            if ($c -eq 4){ $tr.Font.Color.RGB=$cyan; $tr.Font.Bold=[int]1 }
            if ($c -eq 5){ $tr.Font.Color.RGB=$white }
        }
    }
}
$fnR = "Glossary p.40" + $nd + "42 is a reference appendix " + $md + " use on demand, do not present every row."
Add-Box $sRun 40 502 ($W-80) 18 $fnR 8.5 0 $muted $bodyFont 1 | Out-Null

# ===== STEP 7: build HOW-TO-RUN on slide 3 =====
$howTitle = "How to run the room " + $nd + " roles, decision criteria & ownership"
Add-Box $sHow 40 22 ($W-80) 30 $howTitle 22 1 $white $titleFont 1 | Out-Null
Add-Box $sHow ($W-330) 18 250 18 "Future Industrial PLM" 9 0 $muted $bodyFont 3 | Out-Null
Add-Box $sHow ($W-52) 16 38 18 "3" 11 0 $muted $bodyFont 3 | Out-Null

$pTop1=64; $pH1=222; $pTop2=296; $pH2=224
$cLW=40; $cLWd=430; $cRW=490; $cRWd=430

$null = Add-Panel $sHow $cLW $pTop1 $cLWd $pH1 $panel $cyan 1.25
Add-Box $sHow ($cLW+14) ($pTop1+10) ($cLWd-28) 18 "ROLES & PRE-READS" 11 1 $cyan $bodyFont 1 | Out-Null
$rolesTxt = "Chair / decision owner " + $md + " runs the room, holds the decision" + $CR +
            "PLM SME (presenter) " + $md + " walks the narrative" + $CR +
            "Solution Architect " + $md + " architecture & feasibility" + $CR +
            "Development lead " + $md + " feasibility, sequencing, estimates" + $CR +
            "Planning lead " + $md + " position, value, scope, gates" + $CR +
            "Naval architecture / Class " + $md + " solver & evidence acceptance" + $CR +
            "Scribe " + $md + " logs decisions, owners, actions"
Add-Box $sHow ($cLW+14) ($pTop1+34) ($cLWd-28) 120 $rolesTxt 9.5 0 $ltgray $bodyFont 1 | Out-Null
$preTxt = "Pre-reads (24" + $nd + "48h before): p.1" + $nd + "2, p.20" + $nd + "23, p.32" + $nd + "33, p.38; Glossary p.40" + $nd + "42." + $CR +
          "Ground rule (p.4): not a tool-replacement story " + $md + " a governed, auditable, operations-aware decision loop."
Add-Box $sHow ($cLW+14) ($pTop1+158) ($cLWd-28) 56 $preTxt 9 0 $muted $bodyFont 1 | Out-Null

$null = Add-Panel $sHow $cRW $pTop1 $cRWd $pH1 $panel $green 1.25
Add-Box $sHow ($cRW+14) ($pTop1+10) ($cRWd-28) 18 "DECISION CRITERIA" 11 1 $green $bodyFont 1 | Out-Null
$kpi = "contract 100% " + $md + " config " + $ge + "99% " + $md + " impact recall " + $ge + "95% " + $md + " provenance 100% " + $md + " approval " + $le + "5s"
$apTxt = "APPROVE " + $ar + " Phase 2 configuration core + one bounded Continuous Naval Assurance pilot, with named owners; KPI thresholds (" + $kpi + "); assumption boundaries; and one pilot scenario (one hull/block " + $md + " selected loading cases " + $md + " one approved solver adapter " + $md + " seeded change & failure cases)."
Add-Box $sHow ($cRW+14) ($pTop1+34) ($cRWd-28) 116 $apTxt 9.5 0 $ltgray $bodyFont 1 | Out-Null
$ccTxt = "CONDITIONAL " + $ar + " approve with a named backlog to clear before G1/G2." + $CR +
         "DEFER " + $ar + " list the specific blockers + the evidence needed to revisit."
Add-Box $sHow ($cRW+14) ($pTop1+154) ($cRWd-28) 60 $ccTxt 9.5 0 $white $bodyFont 1 | Out-Null

$null = Add-Panel $sHow $cLW $pTop2 $cLWd $pH2 $panel $amber 1.25
Add-Box $sHow ($cLW+14) ($pTop2+10) ($cLWd-28) 18 "DEVELOPMENT OWNS" 11 1 $amber $bodyFont 1 | Out-Null
$buildTxt = "BUILD: OpenAPI contracts (objects, BOM, effectivity, baseline, change impact, evidence); Local Agent / Plugin callbacks from Windows authoring; AI-gate provider interface (rules " + $ar + " RAG " + $ar + " CONNECT); seed truth cases; E2E tests; KPI gates."
Add-Box $sHow ($cLW+14) ($pTop2+34) ($cLWd-28) 108 $buildTxt 9.5 0 $ltgray $bodyFont 1 | Out-Null
$avoidTxt = "AVOID: E3D / Marine / Hull on Linux; big-bang solver replacement; automatic AI release authority; unbounded process customization without approval gates."
Add-Box $sHow ($cLW+14) ($pTop2+146) ($cLWd-28) 68 $avoidTxt 9.5 0 $white $bodyFont 1 | Out-Null

$null = Add-Panel $sHow $cRW $pTop2 $cRWd $pH2 $panel $violet 1.25
Add-Box $sHow ($cRW+14) ($pTop2+10) ($cRWd-28) 18 "PLANNING OWNS  +  CAPTURE LIVE" 11 1 $violet $bodyFont 1 | Out-Null
$planTxt = "PLANNING OWNS: product position (control plane above tools); the customer-value story; 26-week gate discipline; the decision ask; which process variants are allowed."
Add-Box $sHow ($cRW+14) ($pTop2+34) ($cRWd-28) 96 $planTxt 9.5 0 $ltgray $bodyFont 1 | Out-Null
$capTxt = "CAPTURE LIVE: Decision log " + $md + " Action log (action " + $md + " owner " + $md + " due " + $md + " gate) " + $md + " Parking lot for off-scope items (record, don't debate)."
Add-Box $sHow ($cRW+14) ($pTop2+134) ($cRWd-28) 80 $capTxt 9.5 0 $white $bodyFont 1 | Out-Null

# ===== STEP 8: speaker notes for the two new slides =====
try { $sRun.NotesPage.Shapes.Placeholders.Item(2).TextFrame.TextRange.Text = "Open by stating the single decision and the 90-minute path. Keep each block to its time box; the goal is a bounded approval, not a full PLM walkthrough." } catch {}
try { $sHow.NotesPage.Shapes.Placeholders.Item(2).TextFrame.TextRange.Text = "Set expectations: development judges feasibility and sequencing; planning judges position, value and gates. Capture decision, owners, actions and parking-lot items live." } catch {}

# ===== STEP 9: save =====
$pres.Save()
$pres.SaveAs($alias)
$pres.SaveAs($final)
$pres.SaveAs($fpdf, 32)
$cnt = $pres.Slides.Count
$pres.Close()
$pp.Quit()

Write-Output ("SLIDES: " + $cnt)
if ($mismatch.Count -gt 0) { Write-Output "MISMATCHES:"; $mismatch | ForEach-Object { Write-Output ("  " + $_) } } else { Write-Output "RENUMBER: all matched" }
Write-Output ("SAVED V16/alias/final/pdf in " + $dir)
