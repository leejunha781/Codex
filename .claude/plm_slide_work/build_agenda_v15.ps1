$ErrorActionPreference = "Stop"

$base = "D:\이력서\AVEVA - Marine Principal Technical Support & Consultant – PLM SME, Busan\Proposal\"
$src  = $base + "Future_Industrial_PLM_Meeting_Deck_EN_V15.pptx"
$pptx = $base + "Future_Industrial_PLM_Technical_Meeting_Agenda_EN.pptx"
$pdf  = $base + "Future_Industrial_PLM_Technical_Meeting_Agenda_EN.pdf"

# ---- special chars (avoid encoding issues) ----
$nd = [char]0x2013   # en dash
$md = [char]0x00B7   # middot
$ar = [char]0x2192   # arrow
$ge = [char]0x2265   # >=
$le = [char]0x2264   # <=
$CR = [char]13

# ---- palette (OLE BGR ints) ----
$navy   = 3678739     # RGB 19,34,56
$panel  = 4205084     # RGB 28,42,64
$cyan   = 15646772    # RGB 52,192,238
$amber  = 3846888     # RGB 232,178,58
$green  = 9224777     # RGB 73,194,140
$white  = 16777215
$ltgray = 15062727    # #C7D6E5
$muted  = 12626067    # #93A8C0
$violet = 16419751    # RGB 167,139,250
$hdrfill = 2826529    # darker navy RGB 33,34,43 -> just reuse a deep tone

$pp = $null
for ($a=0; $a -lt 6 -and $pp -eq $null; $a++) {
    try { $pp = New-Object -ComObject PowerPoint.Application }
    catch { Start-Sleep -Milliseconds 1200; $pp = $null }
}
if ($pp -eq $null) { throw "PowerPoint COM did not start after retries" }
try { $pp.Visible = $true } catch {}

# read source dimensions
$s = $pp.Presentations.Open($src, $true, $false, $false)
$W = $s.PageSetup.SlideWidth
$H = $s.PageSetup.SlideHeight
$s.Close()

$pres = $pp.Presentations.Add($true)
$pres.PageSetup.SlideWidth  = $W
$pres.PageSetup.SlideHeight = $H

$LAY_BLANK = 12  # ppLayoutBlank

function New-Slide($pres) {
    $idx = $pres.Slides.Count + 1
    $lay = $pres.SlideMaster.CustomLayouts.Item(7)
    $sl = $pres.Slides.AddSlide($idx, $lay)
    foreach ($sh in @($sl.Shapes)) { $sh.Delete() }
    $sl.FollowMasterBackground = $false
    $sl.Background.Fill.Solid()
    $sl.Background.Fill.ForeColor.RGB = $navy
    return $sl
}

function Add-Box($sl,$l,$t,$w,$h,$text,$size,$bold,$color,$font,$align) {
    $tb = $sl.Shapes.AddTextbox(1, [single]$l, [single]$t, [single]$w, [single]$h)
    $tb.TextFrame.WordWrap = -1
    $tb.TextFrame.AutoSize = 0
    $tb.TextFrame.MarginLeft = [single]2
    $tb.TextFrame.MarginRight = [single]2
    $tb.TextFrame.MarginTop = [single]2
    $tb.TextFrame.MarginBottom = [single]2
    $tr = $tb.TextFrame.TextRange
    $tr.Text = $text
    $tr.Font.Size = [single]$size
    $tr.Font.Bold = [int]$bold
    $tr.Font.Name = $font
    $tr.Font.Color.RGB = $color
    $tr.ParagraphFormat.Alignment = [int]$align
    return $tb
}

function Add-Panel($sl,$l,$t,$w,$h,$fill,$line,$lw) {
    $p = $sl.Shapes.AddShape(5, [single]$l, [single]$t, [single]$w, [single]$h)
    $p.Fill.Solid()
    $p.Fill.ForeColor.RGB = $fill
    $p.Line.Visible = -1
    $p.Line.ForeColor.RGB = $line
    $p.Line.Weight = [single]$lw
    $p.Shadow.Visible = 0
    return $p
}

$titleFont = "Aptos Display"
$bodyFont  = "Aptos"

# =================== SLIDE 1 : RUN-OF-SHOW ===================
$s1 = New-Slide $pres

Add-Box $s1 40 22 760 36 "Technical meeting agenda" 26 1 $white $titleFont 1 | Out-Null
$sub1 = "Future Industrial PLM " + $md + " AVEVA Marine " + $md + " Development + Planning decision meeting " + $md + " 90 minutes " + $md + " Deck V15 (41 slides)"
Add-Box $s1 40 60 ($W-80) 22 $sub1 11.5 0 $ltgray $bodyFont 1 | Out-Null
$pgt1 = Add-Box $s1 ($W-80) 18 40 20 "A1" 11 0 $muted $bodyFont 3
$brand1 = Add-Box $s1 ($W-300) 18 250 18 "Future Industrial PLM" 9 0 $muted $bodyFont 3

# decision banner
$ban = Add-Panel $s1 40 92 ($W-80) 50 $panel $cyan 1.5
$banTxt = "DECISION REQUESTED TODAY  " + $md + "  Approve Phase 2 configuration core + one bounded Continuous Naval Assurance pilot " + $ar + " named owners, KPI thresholds, assumption boundaries and the first evidence scenario."
$bt = Add-Box $s1 54 99 ($W-108) 38 $banTxt 11 1 $white $bodyFont 1
$bt.TextFrame.VerticalAnchor = 3

# run-of-show table
$rows = @(
 @("#","Min","Agenda block","Deck","Lead","Outcome to capture / decide"),
 @("0","5","Open & decision objective","p.1","Chair","Confirm the exact approval requested today"),
 @("1","5","How we'll decide " + $nd + " procedure","p.3" + $nd + "5","Chair " + $md + " SME","Agree decision flow; set dev vs planning lenses"),
 @("2","10","Current state & why now","p.6" + $nd + "9","SME","Accept that the gap and urgency are real"),
 @("3","15","Winning ideas / strategy","p.10" + $nd + "18","SME","Agree win-above position + the AI stance"),
 @("4","20","Proposal & architecture","p.19" + $nd + "29","SME " + $md + " Architect","Confirm hybrid architecture is feasible & bounded"),
 @("5","10","Delivery " + $nd + " WBS & KPI gates","p.30" + $nd + "32","SME","Agree the 26-wk staged plan + KPI acceptance"),
 @("6","5","Future-state payoff","p.33","SME","Judge the competitive upside credible"),
 @("7","10","Review " + $nd + " ownership, risks, evidence","p.34" + $nd + "36","Chair","Assign owners; accept risks & assumptions"),
 @("8","10","Decision & close","p.37" + $nd + "38","Chair","Record decision, scope, owners, first scenario")
)

$tLeft = 40; $tTop = 152; $tWidth = ($W-80)
$nRows = $rows.Count; $nCols = 6
$tHeight = 372
$tblShape = $s1.Shapes.AddTable($nRows, $nCols, [single]$tLeft, [single]$tTop, [single]$tWidth, [single]$tHeight)
$tbl = $tblShape.Table

# column widths must sum to $tWidth (=880 if W=960)
$frac = @(0.040, 0.052, 0.262, 0.082, 0.150, 0.414)
for ($c=1; $c -le $nCols; $c++) { $tbl.Columns.Item($c).Width = [single]($tWidth * $frac[$c-1]) }

$aligns = @(2,2,1,2,1,1)  # center,center,left,center,left,left
for ($r=1; $r -le $nRows; $r++) {
    $tbl.Rows.Item($r).Height = [single]34
    for ($c=1; $c -le $nCols; $c++) {
        $cell = $tbl.Cell($r,$c)
        $cell.Shape.TextFrame.VerticalAnchor = 3
        $cell.Shape.TextFrame.MarginLeft = [single]6
        $cell.Shape.TextFrame.MarginRight = [single]6
        $cell.Shape.TextFrame.MarginTop = [single]2
        $cell.Shape.TextFrame.MarginBottom = [single]2
        $tr = $cell.Shape.TextFrame.TextRange
        $tr.Text = [string]$rows[$r-1][$c-1]
        $tr.Font.Name = $bodyFont
        $tr.ParagraphFormat.Alignment = [int]$aligns[$c-1]
        # borders subtle
        for ($b=1; $b -le 4; $b++) {
            try { $cell.Borders.Item($b).ForeColor.RGB = $navy; $cell.Borders.Item($b).Weight=[single]1 } catch {}
        }
        if ($r -eq 1) {
            $cell.Shape.Fill.ForeColor.RGB = $panel
            $tr.Font.Size = [single]10.5
            $tr.Font.Bold = [int]1
            $tr.Font.Color.RGB = $cyan
        } else {
            if ($r % 2 -eq 0) { $cell.Shape.Fill.ForeColor.RGB = $navy } else { $cell.Shape.Fill.ForeColor.RGB = $panel }
            $tr.Font.Size = [single]9
            $tr.Font.Bold = [int]0
            $tr.Font.Color.RGB = $ltgray
            if ($c -eq 1) { $tr.Font.Bold=[int]1; $tr.Font.Color.RGB = $cyan; $tr.Font.Size=[single]10 }
            if ($c -eq 2) { $tr.Font.Color.RGB = $white }
            if ($c -eq 3) { $tr.Font.Bold=[int]1; $tr.Font.Color.RGB = $white; $tr.Font.Size=[single]9.5 }
            if ($c -eq 4) { $tr.Font.Color.RGB = $cyan; $tr.Font.Bold=[int]1 }
            if ($c -eq 5) { $tr.Font.Color.RGB = $white }
        }
    }
}

# footer note
$fn1 = "Glossary p.39" + $nd + "41 is a reference appendix " + $md + " use on demand, do not present every row."
Add-Box $s1 40 ($tTop+$tHeight+6) ($W-80) 18 $fn1 8.5 0 $muted $bodyFont 1 | Out-Null

# =================== SLIDE 2 : SUPPORT ===================
$s2 = New-Slide $pres
Add-Box $s2 40 22 ($W-80) 30 "How to run the room " + $nd + " roles, decision criteria & ownership" 22 1 $white $titleFont 1 | Out-Null
Add-Box $s2 ($W-300) 18 250 18 "Future Industrial PLM" 9 0 $muted $bodyFont 3 | Out-Null
Add-Box $s2 ($W-80) 18 40 20 "A2" 11 0 $muted $bodyFont 3 | Out-Null

# panel geometry
$pTop1 = 64; $pH1 = 222
$pTop2 = 296; $pH2 = 224
$colLW = 40; $colLWd = 430
$colRW = 490; $colRWd = 430

# Panel A : roles & pre-reads
$pa = Add-Panel $s2 $colLW $pTop1 $colLWd $pH1 $panel $cyan 1.25
Add-Box $s2 ($colLW+14) ($pTop1+10) ($colLWd-28) 18 "ROLES & PRE-READS" 11 1 $cyan $bodyFont 1 | Out-Null
$rolesTxt = "Chair / decision owner " + $md + " runs the room, holds the decision" + $CR +
            "PLM SME (presenter) " + $md + " walks the narrative" + $CR +
            "Solution Architect " + $md + " architecture & feasibility" + $CR +
            "Development lead " + $md + " feasibility, sequencing, estimates" + $CR +
            "Planning lead " + $md + " position, value, scope, gates" + $CR +
            "Naval architecture / Class " + $md + " solver & evidence acceptance" + $CR +
            "Scribe " + $md + " logs decisions, owners, actions"
Add-Box $s2 ($colLW+14) ($pTop1+34) ($colLWd-28) 120 $rolesTxt 9.5 0 $ltgray $bodyFont 1 | Out-Null
$preTxt = "Pre-reads (24" + $nd + "48h before): p.1" + $nd + "2, p.19" + $nd + "22, p.31" + $nd + "32, p.37; Glossary p.39" + $nd + "41." + $CR +
          "Ground rule (p.3): not a tool-replacement story " + $md + " a governed, auditable, operations-aware decision loop."
Add-Box $s2 ($colLW+14) ($pTop1+158) ($colLWd-28) 56 $preTxt 9 0 $muted $bodyFont 1 | Out-Null

# Panel B : decision criteria
$pb = Add-Panel $s2 $colRW $pTop1 $colRWd $pH1 $panel $green 1.25
Add-Box $s2 ($colRW+14) ($pTop1+10) ($colRWd-28) 18 "DECISION CRITERIA" 11 1 $green $bodyFont 1 | Out-Null
$kpi = "contract 100% " + $md + " config " + $ge + "99% " + $md + " impact recall " + $ge + "95% " + $md + " provenance 100% " + $md + " approval " + $le + "5s"
$apTxt = "APPROVE " + $ar + " Phase 2 configuration core + one bounded Continuous Naval Assurance pilot, with named owners; KPI thresholds (" + $kpi + "); assumption boundaries; and one pilot scenario (one hull/block " + $md + " selected loading cases " + $md + " one approved solver adapter " + $md + " seeded change & failure cases)."
Add-Box $s2 ($colRW+14) ($pTop1+34) ($colRWd-28) 116 $apTxt 9.5 0 $ltgray $bodyFont 1 | Out-Null
$ccTxt = "CONDITIONAL " + $ar + " approve with a named backlog to clear before G1/G2." + $CR +
         "DEFER " + $ar + " list the specific blockers + the evidence needed to revisit."
Add-Box $s2 ($colRW+14) ($pTop1+154) ($colRWd-28) 60 $ccTxt 9.5 0 $white $bodyFont 1 | Out-Null

# Panel C : development owns
$pc = Add-Panel $s2 $colLW $pTop2 $colLWd $pH2 $panel $amber 1.25
Add-Box $s2 ($colLW+14) ($pTop2+10) ($colLWd-28) 18 "DEVELOPMENT OWNS" 11 1 $amber $bodyFont 1 | Out-Null
$buildTxt = "BUILD: OpenAPI contracts (objects, BOM, effectivity, baseline, change impact, evidence); Local Agent / Plugin callbacks from Windows authoring; AI-gate provider interface (rules " + $ar + " RAG " + $ar + " CONNECT); seed truth cases; E2E tests; KPI gates."
Add-Box $s2 ($colLW+14) ($pTop2+34) ($colLWd-28) 108 $buildTxt 9.5 0 $ltgray $bodyFont 1 | Out-Null
$avoidTxt = "AVOID: E3D / Marine / Hull on Linux; big-bang solver replacement; automatic AI release authority; unbounded process customization without approval gates."
Add-Box $s2 ($colLW+14) ($pTop2+146) ($colLWd-28) 68 $avoidTxt 9.5 0 $white $bodyFont 1 | Out-Null

# Panel D : planning owns + capture
$pd = Add-Panel $s2 $colRW $pTop2 $colRWd $pH2 $panel $violet 1.25
Add-Box $s2 ($colRW+14) ($pTop2+10) ($colRWd-28) 18 "PLANNING OWNS  +  CAPTURE LIVE" 11 1 $violet $bodyFont 1 | Out-Null
$planTxt = "PLANNING OWNS: product position (control plane above tools); the customer-value story; 26-week gate discipline; the decision ask; which process variants are allowed."
Add-Box $s2 ($colRW+14) ($pTop2+34) ($colRWd-28) 96 $planTxt 9.5 0 $ltgray $bodyFont 1 | Out-Null
$capTxt = "CAPTURE LIVE: Decision log " + $md + " Action log (action " + $md + " owner " + $md + " due " + $md + " gate) " + $md + " Parking lot for off-scope items (record, don't debate)."
Add-Box $s2 ($colRW+14) ($pTop2+134) ($colRWd-28) 80 $capTxt 9.5 0 $white $bodyFont 1 | Out-Null

# ---- save ----
if (Test-Path $pptx) { Remove-Item $pptx -Force }
if (Test-Path $pdf)  { Remove-Item $pdf -Force }
$pres.SaveAs($pptx)            # default pptx
$pres.SaveAs($pdf, 32)         # ppSaveAsPDF = 32
$pres.Close()
$pp.Quit()
Write-Output ("SAVED: " + $pptx)
Write-Output ("SAVED: " + $pdf)
