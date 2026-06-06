$ErrorActionPreference = "Stop"

$dir = "D:\이력서\AVEVA - Marine Principal Technical Support & Consultant – PLM SME, Busan\Proposal"
$deck = Join-Path $dir "Future_Industrial_PLM_Meeting_Deck_EN_Rebuilt_v2.pptx"
$backupDir = Join-Path $dir "_backup_20260606"
$reviewDir = "C:\Users\namma\.claude\plm_slide_work\restructure_review_20260606"
New-Item -ItemType Directory -Force -Path $backupDir | Out-Null
New-Item -ItemType Directory -Force -Path $reviewDir | Out-Null
$stamp = Get-Date -Format "yyyyMMdd_HHmmss"
$backup = Join-Path $backupDir ("Future_Industrial_PLM_Meeting_Deck_EN_Rebuilt_v2_before_restructure_" + $stamp + ".pptx")
Copy-Item -LiteralPath $deck -Destination $backup -Force
Write-Output ("BACKUP=" + $backup)

function Rgb([string]$hex) {
    $hex = $hex.TrimStart('#')
    [int](
        [Convert]::ToInt32($hex.Substring(0,2),16) +
        [Convert]::ToInt32($hex.Substring(2,2),16) * 256 +
        [Convert]::ToInt32($hex.Substring(4,2),16) * 65536
    )
}

function Retry($sb) {
    for ($a=0; $a -lt 8; $a++) {
        try { return (& $sb) } catch { Start-Sleep -Milliseconds 100 }
    }
    return (& $sb)
}

$WHITE   = "FFFFFF"
$LIGHT   = "C7D6E5"
$MUTE    = "93A8C0"
$CYAN    = "34C0EE"
$BLUE    = "4F8BE8"
$GREEN   = "49C28C"
$ORANGE  = "F0832F"
$VIOLET  = "9B7BE0"
$MAGENTA = "E653B5"
$AMBER   = "E8B23A"
$RED     = "E5634E"
$GREY    = "8FA9C4"
$BGD     = "0B1626"
$PANEL   = "14253B"
$PANEL2  = "102033"
$ROW     = "101F33"
$LINEC   = "2C4663"

$ppt = $null
$pres = $null
$openedByScript = $true

try {
    $ppt = New-Object -ComObject PowerPoint.Application
    $pres = $ppt.Presentations.Open($deck,0,0,0)
    Write-Output ("COUNT_BEFORE=" + $pres.Slides.Count)

    function NoShadow($s) { try { $s.Shadow.Visible = 0 } catch {} }

    function Panel($l,$t,$w,$h,$fill,$line,$lw,$rad) {
        $s = $script:slide.Shapes.AddShape(5,[single]$l,[single]$t,[single]$w,[single]$h)
        try { $s.Adjustments.Item(1) = [single]$rad } catch {}
        $s.Fill.Solid()
        Retry { $s.Fill.ForeColor.RGB = [int](Rgb $fill) } | Out-Null
        if ($line) {
            $s.Line.Visible = -1
            Retry { $s.Line.ForeColor.RGB = [int](Rgb $line) } | Out-Null
            $s.Line.Weight = [single]$lw
        } else {
            $s.Line.Visible = 0
        }
        NoShadow $s
        return $s
    }

    function Bar($l,$t,$w,$h,$fill) {
        $s = $script:slide.Shapes.AddShape(1,[single]$l,[single]$t,[single]$w,[single]$h)
        $s.Fill.Solid()
        Retry { $s.Fill.ForeColor.RGB = [int](Rgb $fill) } | Out-Null
        $s.Line.Visible = 0
        NoShadow $s
        return $s
    }

    function Oval($l,$t,$w,$h,$fill) {
        $s = $script:slide.Shapes.AddShape(9,[single]$l,[single]$t,[single]$w,[single]$h)
        $s.Fill.Solid()
        Retry { $s.Fill.ForeColor.RGB = [int](Rgb $fill) } | Out-Null
        $s.Line.Visible = 0
        NoShadow $s
        return $s
    }

    function Txt($l,$t,$w,$h,$text,$size,$bold,$color,$font,$align,$anchor) {
        $tb = $script:slide.Shapes.AddTextbox(1,[single]$l,[single]$t,[single]$w,[single]$h)
        $tb.Fill.Visible = 0
        $tb.Line.Visible = 0
        NoShadow $tb
        $tf = $tb.TextFrame
        $tf.WordWrap = -1
        $tf.AutoSize = 0
        $tf.MarginLeft = [single]2
        $tf.MarginRight = [single]2
        $tf.MarginTop = [single]1
        $tf.MarginBottom = [single]1
        $tf.VerticalAnchor = [int]$anchor
        $tr = $tf.TextRange
        $tr.Text = [string]$text
        $tr.Font.Size = [single]$size
        $tr.Font.Bold = [int]$bold
        $tr.Font.Name = [string]$font
        Retry { $tr.Font.Color.RGB = [int](Rgb $color) } | Out-Null
        $tr.ParagraphFormat.Alignment = [int]$align
        try {
            $tr.ParagraphFormat.SpaceBefore = [single]0
            $tr.ParagraphFormat.SpaceAfter = [single]0
        } catch {}
        return $tb
    }

    function Arrow($x1,$y1,$x2,$y2,$color,$w) {
        $c = $script:slide.Shapes.AddConnector(1,[single]$x1,[single]$y1,[single]$x2,[single]$y2)
        Retry { $c.Line.ForeColor.RGB = [int](Rgb $color) } | Out-Null
        $c.Line.Weight = [single]$w
        $c.Line.BeginArrowheadStyle = [int]1
        $c.Line.EndArrowheadStyle = [int]2
        NoShadow $c
        return $c
    }

    function Chip($l,$t,$w,$h,$text,$fill,$line,$textColor,$size) {
        Panel $l $t $w $h $fill $line 0.8 0.30 | Out-Null
        Txt ($l+3) $t ($w-6) $h $text $size 0 $textColor "Aptos" 2 3 | Out-Null
    }

    function DelAll($sl) {
        foreach ($sh in @($sl.Shapes)) { try { $sh.Delete() } catch {} }
    }

    function BuildHeader($title,$subtitle) {
        $bg = $script:slide.Shapes.AddShape(1,[single]0,[single]0,[single]960,[single]540)
        $bg.Line.Visible = 0
        $bg.Fill.Solid()
        Retry { $bg.Fill.ForeColor.RGB = [int](Rgb $BGD) } | Out-Null
        NoShadow $bg
        Txt 34 18 740 32 $title 22 -1 $WHITE "Aptos Display" 1 1 | Out-Null
        Txt 34 54 840 20 $subtitle 10.5 0 $LIGHT "Aptos" 1 1 | Out-Null
        Txt 800 24 130 15 "Future Industrial PLM" 8 0 $MUTE "Aptos" 3 1 | Out-Null
        Bar 34 82 892 1 $LINEC | Out-Null
    }

    function GetSlidesByPrefix($prefix) {
        $result = @()
        foreach ($sl in $pres.Slides) {
            foreach ($sh in $sl.Shapes) {
                try {
                    if ($sh.HasTextFrame -eq -1 -and $sh.TextFrame.HasText -eq -1 -and
                        $sh.TextFrame.TextRange.Text.Trim().StartsWith($prefix)) {
                        $result += $sl
                        break
                    }
                } catch {}
            }
        }
        return $result
    }

    function GetSlide($prefix) {
        $arr = @(GetSlidesByPrefix $prefix)
        if ($arr.Count -eq 0) { throw ("Slide not found: " + $prefix) }
        return $arr[0]
    }

    function ReplacePrefix($sl,$prefix,$newText) {
        foreach ($sh in $sl.Shapes) {
            try {
                if ($sh.HasTextFrame -eq -1 -and $sh.TextFrame.HasText -eq -1 -and
                    $sh.TextFrame.TextRange.Text.Trim().StartsWith($prefix)) {
                    $sh.TextFrame.TextRange.Text = $newText
                    return $sh
                }
            } catch {}
        }
        return $null
    }

    # Capture all slide references before changing titles.
    $sCover        = $pres.Slides.Item(1)
    $sAgenda       = GetSlide "Meeting outcomes"
    $execs         = @(GetSlidesByPrefix "Executive thesis")
    if ($execs.Count -lt 2) { throw "Two Executive thesis slides were expected." }
    $sThesis       = $execs[0]
    $sCurrentTable = $execs[1]
    $sLandscape    = GetSlide "Current PLM landscape"
    $sRoadmap      = GetSlide "Public roadmap signals"
    $sGap          = GetSlide "Strategic gap to fill"
    $sArchitecture = GetSlide "Target architecture concept"
    $sReference    = GetSlide "Executable reference package"
    $sHardware     = GetSlide "Hardware"
    $sNow          = GetSlide "Where we are now"
    $sPhase2       = GetSlide "Phase 2 scope"
    $sRequirement  = GetSlide "Requirement"
    $sFlow         = GetSlide "AI-gated flow"
    $sDrawing      = GetSlide "Design tool"
    $sVersion      = GetSlide "E3D / Hull design"
    $sProcess      = GetSlide "User-Configurable"
    $sGate         = GetSlide "Executable gate flow"
    $sWBS          = GetSlide "Delivery WBS"
    $sKPI          = GetSlide "KPI model"
    $sAbsorb       = GetSlide "Benchmark absorption map"
    $sOperating    = GetSlide "Operating model"
    $sRisk         = GetSlide "Main risks"
    $sEvidence     = GetSlide "Evidence base"
    $sFutureTable  = GetSlide "Shipbuilding design tools"
    $sNapa         = GetSlide "Win above NAPA"
    $sAssuranceAPI = GetSlide "Continuous Naval Assurance Control Plane"
    $sDecision     = GetSlide "Recommended decision"

    # ------------------------------------------------------------------
    # Cover: broaden the story to include the new idea.
    # ------------------------------------------------------------------
    ReplacePrefix $sCover "Reference architecture" "Current benchmark, winning ideas, Continuous Naval Assurance, delivery WBS and KPI gates" | Out-Null
    ReplacePrefix $sCover "Decision expected today" "Decision expected today | Align on the winning product position, Phase 2 + Naval Assurance pilot, delivery governance and KPI-based acceptance." | Out-Null

    # ------------------------------------------------------------------
    # Slide 2: seven-chapter meeting structure.
    # ------------------------------------------------------------------
    DelAll $sAgenda
    $script:slide = $sAgenda
    BuildHeader "Meeting structure - seven decisions in order" "Move from facts to ideas, then validate the delivery plan and final decision"
    $agenda = @(
        @{n="1"; c=$CYAN;   t="CURRENT STATE"; d="Competitor vs AVEVA comparison, PLM position and today's executable baseline"},
        @{n="2"; c=$BLUE;   t="PUBLIC ROADMAP"; d="Official vendor signals and what is confirmed versus directional"},
        @{n="3"; c=$MAGENTA;t="WINNING IDEAS"; d="Meeting thesis: go beyond competitors with Continuous Naval Assurance"},
        @{n="4"; c=$GREEN;  t="PROPOSAL + DELIVERY"; d="Architecture, workflows and pilot scope - then WBS and KPI gates"},
        @{n="5"; c=$ORANGE; t="FUTURE-STATE COMPARE"; d="Re-score AVEVA after the proposed ideas are applied"},
        @{n="6"; c=$AMBER;  t="REVIEW"; d="Ownership, risks, evidence and assumptions requiring agreement"},
        @{n="7"; c=$VIOLET; t="CLOSE"; d="Approve the next increment and the assurance pilot"}
    )
    $y = 102
    foreach ($a in $agenda) {
        Panel 48 $y 864 46 $PANEL $a.c 0.9 0.06 | Out-Null
        Oval 60 ($y+9) 28 28 $a.c | Out-Null
        Txt 60 ($y+9) 28 28 $a.n 11 -1 $WHITE "Aptos" 2 3 | Out-Null
        Txt 102 ($y+7) 210 17 $a.t 9.5 -1 $a.c "Aptos" 1 1 | Out-Null
        Txt 102 ($y+23) 792 17 $a.d 8.5 0 $LIGHT "Aptos" 1 1 | Out-Null
        $y += 52
    }
    Panel 48 474 864 34 "0E2A36" $CYAN 1.0 0.08 | Out-Null
    Txt 60 474 840 34 "Sequencing rule: WBS and KPI appear only after all proposed capabilities and pilot ideas are consolidated." 9.2 -1 $WHITE "Aptos" 1 3 | Out-Null

    # ------------------------------------------------------------------
    # Current-state competitor comparison table (repurpose duplicate).
    # ------------------------------------------------------------------
    DelAll $sCurrentTable
    $script:slide = $sCurrentTable
    BuildHeader "Current-state comparison - no vendor owns the complete loop" "Today's reality: NAPA leads naval architecture; classical PLM leads configuration; AVEVA leads engineering-to-operations context"
    $cx = @(38,218,408,598,788)
    $cw = @(176,186,186,186,136)
    $heads = @("Capability","AVEVA`nE3D/Marine + AIM/PI","Siemens / Dassault`nPLM platforms","Hexagon / CADMATIC`n/ FORAN","NAPA")
    for ($i=0; $i -lt $heads.Count; $i++) {
        $hc = $MUTE
        if ($i -eq 1) { $hc = $CYAN }
        if ($i -eq 4) { $hc = $GREEN }
        Panel $cx[$i] 112 $cw[$i] 46 $PANEL2 $hc 0.8 0.03 | Out-Null
        Txt ($cx[$i]+4) 112 ($cw[$i]-8) 46 $heads[$i] 8.5 -1 $hc "Aptos" 2 3 | Out-Null
    }
    $rows = @(
        @("Configuration / BOM / change","Partial","Strong","Partial","Limited"),
        @("Hull / outfitting / production","Strong","Partial","Strong","Partial"),
        @("Naval architecture / stability","Partial","Limited","Partial","Strong"),
        @("Operations / asset context","Strong","Partial","Strong","Partial"),
        @("Customer-owned API / process","Strong","Partial","Partial","Partial")
    )
    $ry = 164
    foreach ($r in $rows) {
        Panel 38 $ry 886 48 $ROW $LINEC 0.45 0.02 | Out-Null
        Txt 48 $ry 164 48 $r[0] 8.4 -1 $LIGHT "Aptos" 1 3 | Out-Null
        for ($c=1; $c -le 4; $c++) {
            $val = $r[$c]
            $col = $GREY
            if ($val -eq "Strong") { $col = $GREEN }
            elseif ($val -eq "Partial") { $col = $AMBER }
            Txt ($cx[$c]+4) $ry ($cw[$c]-8) 48 $val 9 -1 $col "Aptos" 2 3 | Out-Null
        }
        $ry += 52
    }
    Panel 38 438 886 50 "0E2A24" $GREEN 1.0 0.06 | Out-Null
    Txt 50 438 112 50 "CURRENT GAP" 10 -1 $GREEN "Aptos" 1 3 | Out-Null
    Txt 168 440 744 46 "No current product closes design change -> naval calculation -> evidence -> approval -> operations feedback. That is the opportunity, not replacing every specialist tool." 9.2 -1 $WHITE "Aptos" 1 3 | Out-Null
    Txt 38 500 886 18 "General positioning from official public information; capability varies by version, configuration and implementation." 7 0 $MUTE "Aptos" 1 1 | Out-Null

    # ------------------------------------------------------------------
    # Roadmap comparison: include NAPA's public signals.
    # ------------------------------------------------------------------
    DelAll $sRoadmap
    $script:slide = $sRoadmap
    BuildHeader "Public roadmap comparison - official signals only" "What vendors publicly emphasize, translated into design implications for the proposed AVEVA platform"
    $road = @(
        @{x=38;  y=110; c=$CYAN;   v="AVEVA"; s="Industrial AI in CONNECT, Unified Engineering, PI Data Infrastructure"; i="Opportunity: close design-to-operations feedback and assurance."},
        @{x=346; y=110; c=$BLUE;   v="Siemens"; s="Teamcenter X SaaS and embedded AI across PLM processes"; i="Absorb: cloud-ready configuration and lifecycle automation."},
        @{x=654; y=110; c=$VIOLET; v="Dassault"; s="3DEXPERIENCE cloud, virtual twin and collaborative platform"; i="Absorb: role-based collaboration and virtual-twin validation."},
        @{x=38;  y=274; c=$GREEN;  v="Hexagon"; s="Octave asset lifecycle intelligence and connected operational context"; i="Absorb: AIM, completions, project controls and handover."},
        @{x=346; y=274; c=$ORANGE; v="PTC"; s="Windchill / Codebeamer AI and PLM-ALM traceability"; i="Absorb: requirements, software trace and AI-assisted change."},
        @{x=654; y=274; c=$AMBER;  v="NAPA"; s="Node Network workflows and OCX 3D model-based approval"; i="Integrate: certified naval calculations and approval evidence."}
    )
    foreach ($r in $road) {
        Panel $r.x $r.y 270 146 $PANEL $r.c 1.0 0.05 | Out-Null
        Bar $r.x $r.y 270 5 $r.c | Out-Null
        Txt ($r.x+14) ($r.y+14) 242 20 $r.v 12 -1 $r.c "Aptos Display" 1 1 | Out-Null
        Txt ($r.x+14) ($r.y+41) 242 43 $r.s 8.5 0 $LIGHT "Aptos" 1 1 | Out-Null
        Txt ($r.x+14) ($r.y+91) 242 43 $r.i 8.3 -1 $WHITE "Aptos" 1 1 | Out-Null
    }
    Panel 38 438 886 46 "152747" $BLUE 1.0 0.04 | Out-Null
    Txt 50 438 108 46 "PRINCIPLE" 10 -1 $BLUE "Aptos" 1 3 | Out-Null
    Txt 164 440 748 42 "Treat roadmap items as directional public signals, not guaranteed scope. Build adapters and an assurance graph that can absorb change without rewriting the platform." 9.2 -1 $WHITE "Aptos" 1 3 | Out-Null

    # ------------------------------------------------------------------
    # Thesis and strategic gap: include assurance control plane.
    # ------------------------------------------------------------------
    ReplacePrefix $sThesis "Requirement ->" "Requirement -> Engineering -> Assure -> BOM -> Build -> Commissioning -> Handover -> Operations -> Maintenance -> Change Feedback" | Out-Null
    ReplacePrefix $sThesis "A vendor-neutral Digital Thread Hub" "A vendor-neutral Digital Thread Hub + Continuous Naval Assurance Control Plane with configuration governance, executable APIs, approved-solver evidence and measurable quality gates." | Out-Null
    ReplacePrefix $sGap "A configuration-governed Digital Thread Hub" "A configuration-governed Digital Thread + Continuous Naval Assurance Control Plane that connects engineering changes to approved-solver evidence, approvals and operations feedback." | Out-Null
    ReplacePrefix $sGap "It complements existing PLM" "It complements NAPA/class solvers and existing PLM/CAD/MES tools instead of replacing them first. This lowers adoption and class-acceptance risk while creating visible value early." | Out-Null

    # ------------------------------------------------------------------
    # Absorption map: add NAPA and new control plane.
    # ------------------------------------------------------------------
    DelAll $sAbsorb
    $script:slide = $sAbsorb
    BuildHeader "What to absorb - and where AVEVA must go beyond" "Use proven vendor strengths, then own the decision, evidence and lifecycle-learning layer"
    $src = @(
        @{v="Siemens - PLM rigor"; d="BOM · baseline · effectivity · change governance"; c=$BLUE},
        @{v="Dassault - collaboration"; d="Role-based collaboration · virtual twin · mfg flow"; c=$VIOLET},
        @{v="Hexagon - asset lifecycle"; d="Asset information · completions · operations context"; c=$GREEN},
        @{v="PTC - requirements / ALM"; d="Requirements · SW trace · AI-assisted change"; c=$ORANGE},
        @{v="NAPA - naval assurance"; d="Stability · hydrostatics · Node Network · OCX evidence"; c=$AMBER}
    )
    $sy = 104
    $mids = @()
    foreach ($r in $src) {
        Panel 42 $sy 324 62 $PANEL $r.c 0.9 0.05 | Out-Null
        Bar 42 $sy 5 62 $r.c | Out-Null
        Txt 56 ($sy+9) 296 17 $r.v 9.2 -1 $WHITE "Aptos" 1 1 | Out-Null
        Txt 56 ($sy+29) 296 26 $r.d 7.8 0 $LIGHT "Aptos" 1 1 | Out-Null
        $mids += ($sy+31)
        $sy += 68
    }
    Panel 584 126 340 292 "0E2436" $CYAN 1.4 0.05 | Out-Null
    Bar 584 126 340 6 $CYAN | Out-Null
    Txt 600 140 308 19 "AVEVA-based PLM + Naval Assurance" 12.5 -1 $CYAN "Aptos Display" 1 1 | Out-Null
    Txt 600 165 308 17 "Customer-owned Engineering -> Operations control plane" 8.5 0 $LIGHT "Aptos" 1 1 | Out-Null
    $bul = @(
        "Configuration-governed digital thread",
        "Assurance impact graph across affected scenarios",
        "Federated approved-solver adapters",
        "Physics-guarded AI diagnosis and alternatives",
        "Signed evidence, approvals and operations learning"
    )
    $by = 202
    foreach ($b in $bul) {
        Oval 604 ($by+4) 8 8 $GREEN | Out-Null
        Txt 620 $by 284 18 $b 8.3 0 $LIGHT "Aptos" 1 1 | Out-Null
        $by += 32
    }
    for ($i=0; $i -lt 5; $i++) { Arrow 366 $mids[$i] 580 (190+$i*42) $src[$i].c 1.5 | Out-Null }
    Panel 42 454 882 48 "0E2A24" $GREEN 1.0 0.06 | Out-Null
    Txt 54 454 112 48 "STRATEGY" 10 -1 $GREEN "Aptos" 1 3 | Out-Null
    Txt 172 456 738 44 "Do not copy a competitor or out-solve NAPA. Orchestrate approved solvers inside one executable, configuration-governed, operations-ready assurance loop." 9.2 -1 $WHITE "Aptos" 1 3 | Out-Null

    # ------------------------------------------------------------------
    # Architecture and scope affected by the new idea.
    # ------------------------------------------------------------------
    ReplacePrefix $sArchitecture "A layered, vendor-neutral" "A layered, vendor-neutral Industrial PLM + Continuous Naval Assurance reference platform" | Out-Null
    ReplacePrefix $sArchitecture "React dashboard" "React dashboard, 3D context, assurance margin view, gate report and collaboration views" | Out-Null
    ReplacePrefix $sArchitecture "FastAPI, OpenAPI" "FastAPI, OpenAPI, JWT/OIDC, role-based promote authority and solver-adapter contracts" | Out-Null
    ReplacePrefix $sArchitecture "Digital Thread Layer" "Assurance + Digital Thread Layer" | Out-Null
    ReplacePrefix $sArchitecture "PostgreSQL + relationship graph" "PostgreSQL + assurance graph + rule engine + evaluate/promote flow + evidence lineage" | Out-Null
    ReplacePrefix $sArchitecture "Adapters for E3D" "Adapters for E3D, NAPA/class/CFD/FEM, PLM, ERP, MES, AIM and PI / CONNECT" | Out-Null

    ReplacePrefix $sPhase2 "Phase 2 scope" "Phase 2 + Continuous Naval Assurance pilot scope" | Out-Null
    ReplacePrefix $sPhase2 "Full configuration management" "Full configuration management plus an E3D/Hull change -> affected scenario -> approved-solver evidence pilot" | Out-Null
    ReplacePrefix $sPhase2 "Configuration is not only" "Configuration is the control layer that makes handover, operations and naval-assurance impact analysis reliable without replacing certified solvers." | Out-Null
    ReplacePrefix $sPhase2 "A governed configuration graph" "A governed configuration + assurance graph where each change links to affected objects, loading cases, solver runs, evidence and approvals." | Out-Null
    $minText = "Minimum acceptance`r1)   Multi-level BOM and effectivity return the correct configuration.`r2)   Impact graph finds all seeded affected objects and naval scenarios.`r3)   Approved solver results retain immutable provenance and evidence.`r4)   AI proposes/explains; naval architect and class authority release."
    $minShape = ReplacePrefix $sPhase2 "Minimum acceptance" $minText
    if ($minShape -ne $null) { $minShape.TextFrame.TextRange.Font.Size = [single]10.5 }

    # ------------------------------------------------------------------
    # Delivery WBS: add assurance pilot after all solution ideas.
    # ------------------------------------------------------------------
    DelAll $sWBS
    $script:slide = $sWBS
    BuildHeader "Delivery WBS - consolidated after the full solution scope" "Six KPI-gated packages: configuration first, then Continuous Naval Assurance and shipyard adoption"
    $weeks = 26
    $x0 = 224
    $x1 = 918
    $span = $x1 - $x0
    $weekScale = $span / $weeks
    $gates = @(2,6,12,16,22,26)
    for ($g=0; $g -lt $gates.Count; $g++) {
        $gx = $x0 + $gates[$g] * $weekScale
        Bar $gx 142 1 294 $LINEC | Out-Null
        Txt ($gx-28) 102 56 16 ("G"+($g+1)) 8.5 -1 $AMBER "Aptos" 2 1 | Out-Null
        Txt ($gx-28) 119 56 14 ("wk"+$gates[$g]) 6.8 0 $MUTE "Aptos" 2 1 | Out-Null
        $d = $script:slide.Shapes.AddShape(4,[single]($gx-5),[single]135,[single]10,[single]10)
        $d.Fill.Solid(); Retry { $d.Fill.ForeColor.RGB = [int](Rgb $AMBER) } | Out-Null
        $d.Line.Visible = 0
    }
    Txt ($x0-20) 119 40 14 "start" 6.8 0 $MUTE "Aptos" 2 1 | Out-Null
    Bar $x0 142 1 294 $LINEC | Out-Null
    $wbs = @(
        @{n="WBS 0"; t="Readiness";          s=0;  d=2; c=$CYAN;   note="Backlog · environment · source · roles"},
        @{n="WBS 1"; t="Phase 1 hardening";  s=2;  d=4; c=$BLUE;   note="S6 gates · OIDC · dashboard · CI"},
        @{n="WBS 2"; t="Configuration core"; s=6;  d=6; c=$GREEN;  note="BOM · effectivity · baseline · APIs"},
        @{n="WBS 3"; t="Change impact";      s=12; d=4; c=$ORANGE; note="ECR/ECO/MOC · impact graph · promote"},
        @{n="WBS 4"; t="Naval assurance";    s=16; d=6; c=$MAGENTA;note="E3D/Hull event · solver adapters · evidence"},
        @{n="WBS 5"; t="Pilot adoption";     s=22; d=4; c=$VIOLET; note="Shipyard UAT · class review · performance"}
    )
    $wy = 160
    foreach ($r in $wbs) {
        Txt 40 $wy 166 17 $r.n 8.5 -1 $r.c "Aptos" 1 1 | Out-Null
        Txt 40 ($wy+17) 166 17 $r.t 8.5 0 $WHITE "Aptos" 1 1 | Out-Null
        $bx = $x0 + $r.s * $weekScale
        $bw = $r.d * $weekScale
        Panel $bx ($wy+1) $bw 24 $r.c $r.c 0.8 0.35 | Out-Null
        Txt $bx ($wy+1) $bw 24 ($r.d.ToString()+" wks") 8.5 -1 $WHITE "Aptos" 2 3 | Out-Null
        Txt ($bx+2) ($wy+28) ([math]::Max($bw,180)) 16 $r.note 7.2 0 $MUTE "Aptos" 1 1 | Out-Null
        $wy += 52
    }
    Txt 40 472 884 17 "Gates -> G1 ready | G2 E2E/OIDC | G3 BOM/effectivity 99% | G4 impact recall >=95% | G5 solver provenance + evidence 100% | G6 UAT / response <=5s" 7.8 -1 $AMBER "Aptos" 1 1 | Out-Null
    Panel 48 500 864 28 "0E2A36" $CYAN 0.8 0.05 | Out-Null
    Txt 58 500 844 28 "Control principle: every package ships API contract, seed truth set, E2E test, KPI threshold, evidence and demo scenario." 8.3 -1 $WHITE "Aptos" 1 3 | Out-Null

    # ------------------------------------------------------------------
    # KPI model: include assurance-specific acceptance.
    # ------------------------------------------------------------------
    DelAll $sKPI
    $script:slide = $sKPI
    BuildHeader "KPI model - configuration and naval assurance acceptance" "Objective gates measure correctness, evidence provenance and decision speed - not subjective progress"
    $metrics = @(
        @{x=36;  y=112; c=$CYAN;    t="API contract coverage";       v="100%"; d="All committed features ship with OpenAPI + tests"},
        @{x=266; y=112; c=$CYAN;    t="E2E gate pass rate";          v=">=95%";d="CI validates seed, evaluate, promote and rollback"},
        @{x=496; y=112; c=$GREEN;   t="BOM / effectivity accuracy";  v=">=99%";d="Correct levels, quantities and scoped configuration"},
        @{x=726; y=112; c=$ORANGE;  t="Impact graph recall";         v=">=95%";d="Find all seeded affected objects and naval scenarios"},
        @{x=36;  y=304; c=$MAGENTA; t="Solver-run provenance";       v="100%"; d="Version, input, result and approval source retained"},
        @{x=266; y=304; c=$MAGENTA; t="Evidence-pack completeness";  v="100%"; d="Rule trace, affected objects and approval pack present"},
        @{x=496; y=304; c=$VIOLET;  t="AI diagnostic precision";     v=">=85%";d="Useful ranked explanations without release authority"},
        @{x=726; y=304; c=$BLUE;    t="Dashboard / approval response";v="<=5 s"; d="Status, margin and approval summary under pilot load"}
    )
    foreach ($m in $metrics) {
        Panel $m.x $m.y 202 166 $PANEL $m.c 0.8 0.06 | Out-Null
        Bar $m.x $m.y 202 5 $m.c | Out-Null
        Txt ($m.x+14) ($m.y+15) 174 32 $m.t 8.8 -1 $WHITE "Aptos" 1 1 | Out-Null
        Txt ($m.x+14) ($m.y+52) 174 42 $m.v 21 -1 $m.c "Aptos Display" 2 3 | Out-Null
        Bar ($m.x+14) ($m.y+101) 174 8 "091728" | Out-Null
        Bar ($m.x+14) ($m.y+101) 152 8 $m.c | Out-Null
        Txt ($m.x+14) ($m.y+116) 174 42 $m.d 7.8 0 $LIGHT "Aptos" 1 1 | Out-Null
    }
    Panel 36 486 892 34 "0E2A24" $GREEN 1.0 0.06 | Out-Null
    Txt 50 486 118 34 "GATE DECISION" 9.5 -1 $GREEN "Aptos" 1 3 | Out-Null
    Txt 174 486 740 34 "PASS | CONDITIONAL PASS with backlog | FAIL + remediation sprint. AI never owns the certified release decision." 8.8 -1 $WHITE "Aptos" 1 3 | Out-Null

    # ------------------------------------------------------------------
    # Future-state competitor comparison: correct the shipbuilding page.
    # ------------------------------------------------------------------
    DelAll $sFutureTable
    $script:slide = $sFutureTable
    BuildHeader "Future-state comparison - AVEVA wins by orchestrating the full assurance loop" "Proposed AVEVA + Continuous Naval Assurance versus specialist solvers and traditional platforms"
    $fx = @(36,236,448,614,780)
    $fw = @(196,208,162,162,144)
    $fh = @("Capability","AVEVA + Continuous`nNaval Assurance","NAPA","Classical PLM`nplatforms","Other ship`ndesign tools")
    for ($i=0; $i -lt $fh.Count; $i++) {
        $hc = $MUTE
        if ($i -eq 1) { $hc = $CYAN }
        if ($i -eq 2) { $hc = $GREEN }
        Panel $fx[$i] 108 $fw[$i] 48 $PANEL2 $hc 0.8 0.03 | Out-Null
        Txt ($fx[$i]+4) 108 ($fw[$i]-8) 48 $fh[$i] 8.3 -1 $hc "Aptos" 2 3 | Out-Null
    }
    $frows = @(
        @("Certified stability / hydrostatics","Federated","Strong","Limited","Partial"),
        @("Hull + outfitting + production","Strong","Partial","Partial","Strong"),
        @("Configuration / change / effectivity","Strong","Partial","Strong","Partial"),
        @("Cross-domain assurance impact graph","Strong - new","Partial","Partial","Limited"),
        @("Evidence + approval orchestration","Strong - new","Partial","Strong","Partial"),
        @("Operations calibration / learning","Strong - new","Partial","Partial","Partial"),
        @("Customer-owned API / process","Strong","Partial","Partial","Partial")
    )
    $fy = 160
    foreach ($r in $frows) {
        Panel 36 $fy 888 38 $ROW $LINEC 0.4 0.02 | Out-Null
        Txt 46 $fy 182 38 $r[0] 7.6 -1 $LIGHT "Aptos" 1 3 | Out-Null
        for ($c=1; $c -le 4; $c++) {
            $val = $r[$c]
            $col = $GREY
            if ($val.StartsWith("Strong")) { $col = $GREEN }
            elseif ($val -eq "Federated") { $col = $CYAN }
            elseif ($val -eq "Partial") { $col = $AMBER }
            Txt ($fx[$c]+4) $fy ($fw[$c]-8) 38 $val 7.9 -1 $col "Aptos" 2 3 | Out-Null
        }
        $fy += 41
    }
    Panel 36 458 888 46 "0E2A24" $GREEN 1.0 0.06 | Out-Null
    Txt 48 458 108 46 "TAKEAWAY" 9.5 -1 $GREEN "Aptos" 1 3 | Out-Null
    Txt 162 460 748 42 "Distinctive position = approved specialist solvers + AVEVA lifecycle context + a customer-owned assurance graph. The moat is orchestration and evidence, not proprietary stability math." 8.8 -1 $WHITE "Aptos" 1 3 | Out-Null
    Txt 36 510 888 14 "Future-state concept comparison; proposed capabilities require pilot validation and class / customer agreement." 6.8 0 $MUTE "Aptos" 1 1 | Out-Null

    # ------------------------------------------------------------------
    # Review ownership / operating model.
    # ------------------------------------------------------------------
    DelAll $sOperating
    $script:slide = $sOperating
    BuildHeader "Review ownership and operating model" "Four accountable groups review the complete proposal before any capability is promoted"
    $roles = @(
        @{x=42;  c=$CYAN;    t="Product Planning"; d="Own roadmap, MVP scope, user outcomes, KPI thresholds and acceptance decisions."},
        @{x=270; c=$AMBER;   t="Naval Architecture + Class"; d="Own criteria, solver validation, evidence adequacy and certified release decision."},
        @{x=498; c=$GREEN;   t="Solution Architecture"; d="Own data model, assurance graph, OpenAPI contracts, adapters and non-functional targets."},
        @{x=726; c=$VIOLET;  t="Development Team"; d="Own implementation, automation, security, performance and defect closure."}
    )
    foreach ($r in $roles) {
        Panel $r.x 118 198 188 $PANEL $r.c 1.0 0.05 | Out-Null
        Bar $r.x 118 198 5 $r.c | Out-Null
        Txt ($r.x+12) 136 174 38 $r.t 10 -1 $WHITE "Aptos" 2 3 | Out-Null
        Txt ($r.x+14) 184 170 106 $r.d 8.4 0 $LIGHT "Aptos" 2 1 | Out-Null
        Arrow ($r.x+99) 306 ($r.x+99) 346 $r.c 1.5 | Out-Null
    }
    Panel 98 348 764 132 "0E2436" $CYAN 1.0 0.05 | Out-Null
    Txt 116 362 220 22 "Weekly steering and review" 11.5 -1 $CYAN "Aptos Display" 1 1 | Out-Null
    Txt 116 392 720 74 "1) Review KPI and evidence dashboard.   2) Review failing configuration / assurance gates.   3) Confirm class and naval-architecture issues.   4) Approve scope changes.   5) Promote completed capabilities and update risk burndown." 9.3 -1 $WHITE "Aptos" 1 1 | Out-Null

    # ------------------------------------------------------------------
    # Risks affected by the assurance concept.
    # ------------------------------------------------------------------
    DelAll $sRisk
    $script:slide = $sRisk
    BuildHeader "Key review risks and mitigations" "Continuous Naval Assurance adds value only when certified authority, provenance and scope boundaries stay explicit"
    $gx=128; $gy=128; $cw=92; $ch=78
    $zoneF = @(@("3A2E10","3A1614","3A1614"),@("12352A","3A2E10","3A1614"),@("12352A","12352A","3A2E10"))
    $zoneL = @(@($AMBER,$RED,$RED),@($GREEN,$AMBER,$RED),@($GREEN,$GREEN,$AMBER))
    for($r=0;$r -lt 3;$r++){ for($c=0;$c -lt 3;$c++){ Panel ($gx+$c*$cw) ($gy+$r*$ch) ($cw-4) ($ch-4) $zoneF[$r][$c] $zoneL[$r][$c] 0.75 0.04 | Out-Null } }
    $rl=@("High","Med","Low")
    for($r=0;$r -lt 3;$r++){ Txt 78 ($gy+$r*$ch+27) 46 16 $rl[$r] 8.5 -1 $MUTE "Aptos" 3 1 | Out-Null }
    $cl=@("Low","Med","High")
    for($c=0;$c -lt 3;$c++){ Txt ($gx+$c*$cw) ($gy+3*$ch+4) ($cw-4) 14 $cl[$c] 8.5 -1 $MUTE "Aptos" 2 1 | Out-Null }
    Txt 128 ($gy+3*$ch+22) 276 14 "Impact ->" 9 -1 $LIGHT "Aptos" 2 1 | Out-Null
    Txt 40 108 90 14 "Likelihood" 9 -1 $LIGHT "Aptos" 1 1 | Out-Null
    function RiskDot($r,$c,$ox,$num,$col) {
        $x=$gx+$c*$cw+($cw-4)/2-11+$ox
        $y=$gy+$r*$ch+($ch-4)/2-11
        Oval $x $y 22 22 $col | Out-Null
        Txt $x $y 22 22 ([string]$num) 11 -1 $WHITE "Aptos" 2 3 | Out-Null
    }
    RiskDot 0 2 0 1 $RED
    RiskDot 1 2 0 2 $RED
    RiskDot 1 1 -13 3 $AMBER
    RiskDot 1 1 13 4 $AMBER
    RiskDot 2 2 0 5 $AMBER
    $risks = @(
        @{n=1;t="Scope creep";m="Pilot one E3D/Hull scenario; adapters orchestrate solvers instead of rewriting them.";c=$RED},
        @{n=2;t="Solver / class acceptance";m="Certified solver remains source of truth; immutable provenance and human/class release.";c=$RED},
        @{n=3;t="Data and model mapping";m="Seeded truth cases validate hull, compartment, weight, loading case and rule links.";c=$AMBER},
        @{n=4;t="AI overreach / security";m="AI proposes and explains only; OIDC/RBAC, signed evidence and full audit trail.";c=$AMBER},
        @{n=5;t="Adoption and ownership";m="Naval architect, engineering, IT and planning jointly own pilot and KPI decisions.";c=$AMBER}
    )
    Txt 448 112 480 18 "Risk -> mitigation" 10.5 -1 $WHITE "Aptos" 1 1 | Out-Null
    $ly=138
    foreach($rk in $risks){
        Panel 448 $ly 480 62 $PANEL $rk.c 0.75 0.06 | Out-Null
        Oval 460 ($ly+10) 22 22 $rk.c | Out-Null
        Txt 460 ($ly+10) 22 22 ([string]$rk.n) 11 -1 $WHITE "Aptos" 2 3 | Out-Null
        Txt 492 ($ly+7) 424 17 $rk.t 9.5 -1 $WHITE "Aptos" 1 1 | Out-Null
        Txt 492 ($ly+27) 424 28 $rk.m 7.8 0 $LIGHT "Aptos" 1 1 | Out-Null
        $ly += 68
    }

    # ------------------------------------------------------------------
    # Evidence base: include NAPA and separate facts from concepts.
    # ------------------------------------------------------------------
    DelAll $sEvidence
    $script:slide = $sEvidence
    BuildHeader "Evidence base and assumption boundaries" "Official public signals support the direction; proposed control-plane capabilities still require pilot validation"
    $ev = @(
        @{x=38;y=112;c=$CYAN;v="AVEVA";d="CONNECT industrial AI, Unified Engineering, E3D Design and PI System lifecycle context."},
        @{x=346;y=112;c=$BLUE;v="Siemens";d="Teamcenter X SaaS, lifecycle single source of data and embedded AI positioning."},
        @{x=654;y=112;c=$VIOLET;v="Dassault";d="3DEXPERIENCE cloud, collaborative platform and virtual-twin direction."},
        @{x=38;y=256;c=$GREEN;v="Hexagon";d="Octave asset lifecycle intelligence, connected operational context and project controls."},
        @{x=346;y=256;c=$ORANGE;v="PTC";d="Windchill / Codebeamer AI, requirements and PLM-ALM traceability direction."},
        @{x=654;y=256;c=$AMBER;v="NAPA";d="Naval architecture and stability, Node Network workflows and OCX model-based approval."}
    )
    foreach ($e in $ev) {
        Panel $e.x $e.y 270 126 $PANEL $e.c 0.8 0.05 | Out-Null
        Bar $e.x $e.y 270 5 $e.c | Out-Null
        Txt ($e.x+14) ($e.y+15) 242 20 $e.v 11 -1 $e.c "Aptos Display" 1 1 | Out-Null
        Txt ($e.x+14) ($e.y+44) 242 66 $e.d 8.4 0 $LIGHT "Aptos" 1 1 | Out-Null
    }
    Panel 38 406 886 82 "2A2410" $AMBER 1.0 0.06 | Out-Null
    Txt 52 416 132 20 "ASSUMPTION BOUNDARY" 9.5 -1 $AMBER "Aptos" 1 1 | Out-Null
    Txt 190 414 720 58 "Public roadmap signals are not product commitments. The assurance graph, solver adapters, AI diagnosis and closed-loop calibration are proposed AVEVA-based PLM extensions; validate them with HD Hyundai and class authorities in the pilot." 8.8 -1 $WHITE "Aptos" 1 1 | Out-Null
    Txt 38 500 886 18 "Key sources: aveva.com | napa.fi/software-and-services/ship-design/software/napa/ | napa.fi/napa-release-2025-2/ | napa.fi/extending-ocx/ | official vendor product pages" 6.5 0 $MUTE "Aptos" 1 1 | Out-Null

    # Final decision already includes the pilot; sharpen the review language.
    ReplacePrefix $sDecision "Proceed with Phase 2" "Proceed with Phase 2 full configuration management and a bounded Continuous Naval Assurance pilot as the next delivery increment.`r`rApprove: OpenAPI-first development, seeded E2E tests, OIDC security, KPI-gated WBS, certified-solver adapters and an E3D/Hull assurance scenario." | Out-Null
    ReplacePrefix $sDecision "Why now" "Why now`rThe executable reference is ready. The next differentiator is controlled configuration, cross-domain change impact and continuous naval assurance - without replacing certified specialist solvers." | Out-Null

    # ------------------------------------------------------------------
    # Reorder all 28 slides to the requested seven-part narrative.
    # ------------------------------------------------------------------
    $order = @(
        $sCover,
        $sAgenda,
        $sCurrentTable,
        $sLandscape,
        $sNow,
        $sRoadmap,
        $sThesis,
        $sGap,
        $sAbsorb,
        $sNapa,
        $sArchitecture,
        $sReference,
        $sHardware,
        $sPhase2,
        $sRequirement,
        $sFlow,
        $sDrawing,
        $sVersion,
        $sAssuranceAPI,
        $sProcess,
        $sGate,
        $sWBS,
        $sKPI,
        $sFutureTable,
        $sOperating,
        $sRisk,
        $sEvidence,
        $sDecision
    )
    if ($order.Count -ne 28) { throw ("Unexpected order count: " + $order.Count) }
    for ($i=0; $i -lt $order.Count; $i++) { $order[$i].MoveTo([int]($i+1)) }

    # Renumber small page-number shapes in the top-right.
    foreach ($sl in $pres.Slides) {
        $updated = $false
        foreach ($sh in $sl.Shapes) {
            try {
                if ($sh.HasTextFrame -eq -1 -and $sh.TextFrame.HasText -eq -1 -and
                    $sh.Left -gt 890 -and $sh.Top -lt 55 -and $sh.Width -lt 60 -and $sh.Height -lt 35) {
                    $sh.TextFrame.TextRange.Text = [string]$sl.SlideIndex
                    $updated = $true
                }
            } catch {}
        }
        if (-not $updated) {
            $script:slide = $sl
            Txt 905 24 30 16 ([string]$sl.SlideIndex) 8 0 $MUTE "Aptos" 3 1 | Out-Null
        }
    }

    $pres.Save()
    foreach ($sl in $pres.Slides) {
        $sl.Export((Join-Path $reviewDir ("slide-{0:D2}.png" -f $sl.SlideIndex)),"PNG",1280,720)
    }
    Write-Output ("COUNT_AFTER=" + $pres.Slides.Count)
    Write-Output "RESTRUCTURE_DONE"
}
finally {
    if ($pres -ne $null) {
        try { $pres.Close() } catch {}
        try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($pres) | Out-Null } catch {}
    }
    if ($ppt -ne $null) {
        try { if ($ppt.Presentations.Count -eq 0) { $ppt.Quit() } } catch {}
        try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ppt) | Out-Null } catch {}
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
