$ErrorActionPreference = "Stop"

$dir = "D:\이력서\AVEVA - Marine Principal Technical Support & Consultant – PLM SME, Busan\Proposal"
$deck = Join-Path $dir "Future_Industrial_PLM_Meeting_Deck_EN_Rebuilt_v2.pptx"
$backupDir = Join-Path $dir "_backup_20260606"
$reviewDir = "C:\Users\namma\.claude\plm_slide_work\enhanced_full_deck_review_20260606"
New-Item -ItemType Directory -Force -Path $backupDir | Out-Null
New-Item -ItemType Directory -Force -Path $reviewDir | Out-Null
Get-ChildItem -LiteralPath $reviewDir -Filter "slide-*.png" -ErrorAction SilentlyContinue | Remove-Item -Force

$stamp = Get-Date -Format "yyyyMMdd_HHmmss"
$backup = Join-Path $backupDir ("Future_Industrial_PLM_Meeting_Deck_EN_Rebuilt_v2_before_absorb_full_enhancement_" + $stamp + ".pptx")
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
        try { return (& $sb) } catch { Start-Sleep -Milliseconds 120 }
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
$BGD     = "0B1626"
$PANEL   = "14253B"
$PANEL2  = "102033"
$ROW     = "101F33"
$LINEC   = "2C4663"

$ppt = $null
$pres = $null

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
        Panel $l $t $w $h $fill $line 0.8 0.28 | Out-Null
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
        Txt 34 18 790 32 $title 21.5 -1 $WHITE "Aptos Display" 1 1 | Out-Null
        Txt 34 54 880 20 $subtitle 9.6 0 $LIGHT "Aptos" 1 1 | Out-Null
        Txt 800 24 130 15 "Future Industrial PLM" 7.6 0 $MUTE "Aptos" 3 1 | Out-Null
        Bar 34 82 892 1 $LINEC | Out-Null
    }

    function GetSlide($prefix) {
        foreach ($sl in $pres.Slides) {
            foreach ($sh in $sl.Shapes) {
                try {
                    if ($sh.HasTextFrame -eq -1 -and $sh.TextFrame.HasText -eq -1 -and
                        $sh.TextFrame.TextRange.Text.Trim().StartsWith($prefix)) {
                        return $sl
                    }
                } catch {}
            }
        }
        throw ("Slide not found: " + $prefix)
    }

    function ReplacePrefix($sl,$prefix,$newText) {
        foreach ($sh in $sl.Shapes) {
            try {
                if ($sh.HasTextFrame -eq -1 -and $sh.TextFrame.HasText -eq -1 -and
                    $sh.TextFrame.TextRange.Text.Trim().StartsWith($prefix)) {
                    $sh.TextFrame.TextRange.Text = [string]$newText
                    return $sh
                }
            } catch {}
        }
        return $null
    }

    # ------------------------------------------------------------------
    # Slide 8: turn the strategic gap into a competitive control-plane map.
    # ------------------------------------------------------------------
    $sGap = GetSlide "Strategic gap to fill"
    DelAll $sGap
    $script:slide = $sGap
    BuildHeader "Strategic gap to fill" "Competitors own strong domains; nobody owns the open change-to-evidence-to-operations decision loop"

    Txt 42 98 238 18 "SPECIALIST STRENGTHS TO FEDERATE" 8.5 -1 $MUTE "Aptos" 1 1 | Out-Null
    $strengths = @(
        @{y=126;c=$BLUE;t="PRODUCT CONFIGURATION";d="Siemens / Dassault`rBOM · variants · baseline · change · collaboration"},
        @{y=226;c=$AMBER;t="NAVAL ASSURANCE";d="NAPA / class solvers`rHydrostatics · stability · rules · approved calculations"},
        @{y=326;c=$GREEN;t="ASSET LIFECYCLE";d="Hexagon / AVEVA`rHandover · completions · asset and operations context"}
    )
    foreach ($r in $strengths) {
        Panel 42 $r.y 244 82 $PANEL $r.c 0.9 0.05 | Out-Null
        Bar 42 $r.y 5 82 $r.c | Out-Null
        Txt 58 ($r.y+11) 210 17 $r.t 8.8 -1 $r.c "Aptos" 1 1 | Out-Null
        Txt 58 ($r.y+34) 210 39 $r.d 8.0 0 $LIGHT "Aptos" 1 1 | Out-Null
        Arrow 286 ($r.y+41) 326 267 $r.c 1.5 | Out-Null
    }

    Panel 326 104 310 324 "0E2436" $CYAN 1.4 0.05 | Out-Null
    Bar 326 104 310 6 $CYAN | Out-Null
    Txt 344 120 274 18 "THE MISSING PRODUCT" 9.5 -1 $CYAN "Aptos" 2 1 | Out-Null
    Txt 344 145 274 38 "Open AVEVA Decision Control Plane" 15 -1 $WHITE "Aptos Display" 2 3 | Out-Null
    Oval 426 222 110 88 "0C3040" | Out-Null
    Txt 438 236 86 18 "CHANGE ->" 8.5 -1 $CYAN "Aptos" 2 1 | Out-Null
    Txt 438 255 86 34 "DECISION`rLOOP" 10.5 -1 $WHITE "Aptos" 2 1 | Out-Null
    $loopNodes = @(
        @{x=420;y=184;w=122;t="Configuration";c=$BLUE},
        @{x=510;y=213;w=105;t="Impact";c=$ORANGE},
        @{x=502;y=326;w=113;t="Evidence";c=$AMBER},
        @{x=347;y=326;w=113;t="Approval";c=$VIOLET},
        @{x=347;y=213;w=105;t="Operations";c=$GREEN}
    )
    foreach ($n in $loopNodes) { Chip $n.x $n.y $n.w 27 $n.t $PANEL2 $n.c $LIGHT 7.3 }
    Arrow 481 211 548 229 $BLUE 1.2 | Out-Null
    Arrow 562 240 558 323 $ORANGE 1.2 | Out-Null
    Arrow 505 340 462 340 $AMBER 1.2 | Out-Null
    Arrow 404 327 389 241 $VIOLET 1.2 | Out-Null
    Arrow 402 227 448 203 $GREEN 1.2 | Out-Null
    Chip 414 381 134 25 "policy + physics-guarded AI" "2A1431" $MAGENTA $MAGENTA 6.8
    Txt 344 411 274 14 "Own the decision, evidence and learning layer." 7.4 0 $MUTE "Aptos" 2 1 | Out-Null

    Panel 676 104 248 324 $PANEL $GREEN 1.0 0.05 | Out-Null
    Bar 676 104 248 6 $GREEN | Out-Null
    Txt 692 120 216 18 "WHY AVEVA CAN OWN IT" 9.5 -1 $GREEN "Aptos" 1 1 | Out-Null
    Txt 692 145 216 34 "Engineering truth + lifecycle context" 12.0 -1 $WHITE "Aptos Display" 1 1 | Out-Null
    $aveva = @(
        "E3D / Marine design-change truth",
        "Unified Engineering + AIM context",
        "PI / CONNECT live operational truth",
        "Open, vendor-neutral integration",
        "Customer-owned API and process"
    )
    $ay = 204
    foreach ($a in $aveva) {
        Oval 696 ($ay+4) 8 8 $GREEN | Out-Null
        Txt 712 $ay 194 19 $a 7.8 0 $LIGHT "Aptos" 1 1 | Out-Null
        $ay += 34
    }
    Panel 692 374 216 36 "0E2A24" $GREEN 0.8 0.05 | Out-Null
    Txt 700 374 200 36 "Outcome: safer, faster change with visible accountability." 7.8 -1 $WHITE "Aptos" 1 3 | Out-Null

    Panel 42 452 882 55 "0E2A36" $CYAN 1.0 0.05 | Out-Null
    Txt 58 462 138 18 "PRODUCT BOUNDARY" 9 -1 $CYAN "Aptos" 1 1 | Out-Null
    Txt 202 457 706 43 "Federate specialist tools instead of replacing them. Own the cross-domain configuration, impact, evidence, approval and operations-feedback loop." 8.7 -1 $WHITE "Aptos" 1 3 | Out-Null

    # ------------------------------------------------------------------
    # Slide 9: company-by-company absorption map, reflecting the attachment.
    # ------------------------------------------------------------------
    $sAbsorb = GetSlide "What to absorb"
    DelAll $sAbsorb
    $script:slide = $sAbsorb
    BuildHeader "What to absorb - and where AVEVA must go beyond" "Absorb proven strengths company by company; make them executable through AVEVA engineering-to-operations context"

    Txt 38 96 122 16 "COMPANY / PLATFORM" 7.2 -1 $MUTE "Aptos" 1 1 | Out-Null
    Txt 168 96 190 16 "ABSORB THE BEST" 7.2 -1 $MUTE "Aptos" 1 1 | Out-Null
    Txt 364 96 166 16 "AVEVA MOVE BEYOND" 7.2 -1 $MUTE "Aptos" 1 1 | Out-Null
    $vendors = @(
        @{y=116;c=$BLUE;v="SIEMENS`rTeamcenter";a="BOM · variants · baselines`rchange · lifecycle automation";m="Open promote APIs +`roperations-aware configuration"},
        @{y=181;c=$VIOLET;v="DASSAULT`r3DEXPERIENCE / ENOVIA";a="Role collaboration · virtual twin`rmodel-based engineering-to-mfg";m="Twin action -> auditable`rAPI event + evidence"},
        @{y=246;c=$GREEN;v="HEXAGON`rOctave / ALI";a="Asset info · completions`rproject controls · operations";m="Link asset state to`rconfiguration + approval"},
        @{y=311;c=$ORANGE;v="PTC`rCodebeamer / Windchill";a="Requirements · risk · test`rend-to-end traceability + AI";m="Tie requirements to hull,`rBOM, solver and approval"},
        @{y=376;c=$AMBER;v="NAPA`rNaval architecture";a="Hydrostatics · stability · rules`rcalculation workflows";m="Keep solver of record; own`rcontinuous assurance loop"}
    )
    foreach ($r in $vendors) {
        Panel 38 $r.y 494 58 $PANEL $r.c 0.8 0.04 | Out-Null
        Bar 38 $r.y 5 58 $r.c | Out-Null
        Txt 52 ($r.y+8) 108 43 $r.v 7.5 -1 $r.c "Aptos" 1 1 | Out-Null
        Bar 162 ($r.y+7) 1 44 $LINEC | Out-Null
        Txt 170 ($r.y+8) 184 43 $r.a 7.2 0 $LIGHT "Aptos" 1 1 | Out-Null
        Bar 358 ($r.y+7) 1 44 $LINEC | Out-Null
        Txt 366 ($r.y+8) 156 43 $r.m 7.2 -1 $WHITE "Aptos" 1 1 | Out-Null
        Arrow 532 ($r.y+29) 554 ($r.y+29) $r.c 1.4 | Out-Null
    }

    Panel 558 116 366 318 "0E2436" $CYAN 1.4 0.05 | Out-Null
    Bar 558 116 366 7 $CYAN | Out-Null
    Txt 578 132 326 19 "NEW AVEVA-BASED PLM" 11 -1 $CYAN "Aptos Display" 1 1 | Out-Null
    Txt 578 157 326 22 "Engineering -> Operations digital thread" 10.0 0 $LIGHT "Aptos" 1 1 | Out-Null
    Txt 578 190 326 17 "ABSORBS THE BEST OF EACH" 8.6 -1 $MUTE "Aptos" 1 1 | Out-Null
    $best = @(
        "Configuration rigor  |  Siemens-grade",
        "Collaboration + virtual twin  |  Dassault-grade",
        "Asset lifecycle + completions  |  Hexagon-grade",
        "Requirements / ALM trace  |  PTC-grade",
        "Naval calculations + rules  |  NAPA-grade"
    )
    $by = 215
    foreach ($b in $best) {
        Oval 582 ($by+4) 8 8 $GREEN | Out-Null
        Txt 598 $by 300 18 $b 7.7 0 $LIGHT "Aptos" 1 1 | Out-Null
        $by += 27
    }
    Bar 578 352 326 1 $LINEC | Out-Null
    Txt 578 365 326 18 "AVEVA DIFFERENTIATION" 8.8 -1 $CYAN "Aptos" 1 1 | Out-Null
    Txt 578 388 326 39 "E3D / Marine + Unified Engineering + AIM + PI + CONNECT`r+ customer-owned OpenAPI + configuration / assurance graph + policy / AI gates" 7.7 -1 $WHITE "Aptos" 1 1 | Out-Null

    Panel 38 452 886 48 "0E2A24" $GREEN 1.0 0.05 | Out-Null
    Txt 54 452 112 48 "OUTCOME" 9.5 -1 $GREEN "Aptos" 1 3 | Out-Null
    Txt 172 454 738 44 "Operations-ready, configuration-governed, approval-evidenced PLM - open enough to absorb the best specialist capability without surrendering the decision loop." 8.6 -1 $WHITE "Aptos" 1 3 | Out-Null
    Txt 40 510 884 13 "Public product-positioning basis checked 6 Jun 2026; capability details and integration boundaries require customer / vendor validation." 6.1 0 $MUTE "Aptos" 1 1 | Out-Null

    # ------------------------------------------------------------------
    # Slide 16: convert pilot scope into a bounded executable workflow.
    # ------------------------------------------------------------------
    $sPilot = GetSlide "Phase 2 + Continuous Naval Assurance pilot scope"
    DelAll $sPilot
    $script:slide = $sPilot
    BuildHeader "Phase 2 + Continuous Naval Assurance pilot scope" "A bounded executable slice proves configuration, impact, approved-solver evidence, approval authority and operations feedback"

    $steps = @(
        @{x=38;c=$BLUE;n="1";t="E3D / Hull`rchange";api="POST /events"},
        @{x=188;c=$CYAN;n="2";t="Resolve`rconfiguration";api="GET /configuration"},
        @{x=338;c=$ORANGE;n="3";t="Evaluate`rimpact";api="POST /impact/evaluate"},
        @{x=488;c=$AMBER;n="4";t="Run approved`rsolver";api="POST /solver-runs"},
        @{x=638;c=$VIOLET;n="5";t="Evidence +`rapproval";api="POST /promote"},
        @{x=788;c=$GREEN;n="6";t="Operations`rfeedback";api="AIM + PI / CONNECT"}
    )
    foreach ($s in $steps) {
        Panel $s.x 108 132 79 $PANEL $s.c 0.9 0.05 | Out-Null
        Oval ($s.x+10) 119 23 23 $s.c | Out-Null
        Txt ($s.x+10) 119 23 23 $s.n 8.5 -1 $WHITE "Aptos" 2 3 | Out-Null
        Txt ($s.x+39) 116 83 37 $s.t 8.4 -1 $WHITE "Aptos" 1 1 | Out-Null
        Txt ($s.x+10) 159 112 17 $s.api 6.6 0 $s.c "Aptos" 2 1 | Out-Null
    }
    for ($i=0; $i -lt 5; $i++) { Arrow ($steps[$i].x+132) 148 ($steps[$i+1].x-3) 148 $CYAN 1.2 | Out-Null }

    $grade = @(
        @{x=188;w=132;t="Siemens-grade`rconfiguration";c=$BLUE},
        @{x=338;w=132;t="PTC-grade`rtraceability";c=$ORANGE},
        @{x=488;w=132;t="NAPA-grade`rcalculation";c=$AMBER},
        @{x=638;w=132;t="Dassault-grade`rdecision context";c=$VIOLET},
        @{x=788;w=132;t="AVEVA lifecycle`rcontext";c=$GREEN}
    )
    foreach ($g in $grade) { Chip $g.x 197 $g.w 30 $g.t $PANEL2 $g.c $LIGHT 6.7 }

    $cards = @(
        @{x=38;c=$BLUE;t="BOUNDED SCENARIO";d="One hull / block`rSelected loading cases`rOne approved solver adapter`rSeeded change and failure cases"},
        @{x=342;c=$CYAN;t="DELIVERABLE PACKAGE";d="Configuration + assurance graph`rOpenAPI adapter contracts`rStatus / margin dashboard`rSigned approval evidence pack"},
        @{x=646;c=$GREEN;t="RELEASE GOVERNANCE";d="AI proposes, explains and holds`rNaval architect verifies`rApprover / class authority releases`rImmutable provenance retained"}
    )
    foreach ($c in $cards) {
        Panel $c.x 248 276 159 $PANEL $c.c 0.9 0.05 | Out-Null
        Bar $c.x 248 276 5 $c.c | Out-Null
        Txt ($c.x+16) 264 244 18 $c.t 9.0 -1 $c.c "Aptos" 1 1 | Out-Null
        Txt ($c.x+16) 295 244 98 $c.d 8.1 0 $LIGHT "Aptos" 1 1 | Out-Null
    }

    Panel 38 432 884 67 "0E2A36" $CYAN 1.0 0.05 | Out-Null
    Txt 54 441 118 18 "MINIMUM ACCEPTANCE" 8.7 -1 $CYAN "Aptos" 1 1 | Out-Null
    Chip 180 444 168 34 "Configuration accuracy`r>=99%" $PANEL2 $GREEN $WHITE 7.2
    Chip 360 444 168 34 "Impact recall`r>=95%" $PANEL2 $ORANGE $WHITE 7.2
    Chip 540 444 168 34 "Solver provenance`r100%" $PANEL2 $MAGENTA $WHITE 7.2
    Chip 720 444 184 34 "Dashboard / approval`r<=5 s" $PANEL2 $BLUE $WHITE 7.2
    Txt 180 483 724 12 "Release rule: AI never owns approval authority; the customer and class authority remain accountable." 6.7 0 $MUTE "Aptos" 1 1 | Out-Null

    # ------------------------------------------------------------------
    # Slide 30: strengthen the closing decision with scope and timeline.
    # ------------------------------------------------------------------
    $sDecision = GetSlide "Recommended decision"
    DelAll $sDecision
    $script:slide = $sDecision
    BuildHeader "Recommended decision" "Approve the bounded next increment that proves the open AVEVA control plane in a real shipbuilding change scenario"

    Panel 38 110 310 340 "0E2436" $CYAN 1.4 0.05 | Out-Null
    Bar 38 110 310 7 $CYAN | Out-Null
    Txt 58 130 270 20 "APPROVE NOW" 12 -1 $CYAN "Aptos Display" 1 1 | Out-Null
    Txt 58 158 270 46 "Phase 2 configuration core + bounded Continuous Naval Assurance pilot" 13 -1 $WHITE "Aptos Display" 1 1 | Out-Null
    $decisions = @(
        @{n="1";t="PRODUCT POSITION";d="Open AVEVA control plane above specialist tools"},
        @{n="2";t="DELIVERY INCREMENT";d="One E3D/Hull change-to-evidence pilot slice"},
        @{n="3";t="GOVERNANCE";d="26-week WBS, KPI gates, customer + class release"}
    )
    $dy = 230
    foreach ($d in $decisions) {
        Oval 58 $dy 28 28 $CYAN | Out-Null
        Txt 58 $dy 28 28 $d.n 9.2 -1 $WHITE "Aptos" 2 3 | Out-Null
        Txt 98 ($dy-1) 226 16 $d.t 7.8 -1 $CYAN "Aptos" 1 1 | Out-Null
        Txt 98 ($dy+17) 226 28 $d.d 7.5 0 $LIGHT "Aptos" 1 1 | Out-Null
        $dy += 66
    }

    Txt 382 108 542 20 "WHAT APPROVAL UNLOCKS" 9.5 -1 $MUTE "Aptos" 1 1 | Out-Null
    $phases = @(
        @{x=382;c=$BLUE;w="0-6 WEEKS";t="Harden reference";d="OIDC · seed truth · E2E gates"},
        @{x=558;c=$GREEN;w="6-16 WEEKS";t="Configuration + impact";d="BOM · effectivity · baseline · graph"},
        @{x=734;c=$MAGENTA;w="16-26 WEEKS";t="Assurance + UAT";d="Solver evidence · approvals · class review"}
    )
    foreach ($p in $phases) {
        Panel $p.x 140 160 118 $PANEL $p.c 0.9 0.05 | Out-Null
        Bar $p.x 140 160 5 $p.c | Out-Null
        Txt ($p.x+12) 155 136 16 $p.w 7.5 -1 $p.c "Aptos" 1 1 | Out-Null
        Txt ($p.x+12) 181 136 22 $p.t 9.2 -1 $WHITE "Aptos" 1 1 | Out-Null
        Txt ($p.x+12) 213 136 30 $p.d 7.2 0 $LIGHT "Aptos" 1 1 | Out-Null
    }
    Arrow 542 199 554 199 $CYAN 1.2 | Out-Null
    Arrow 718 199 730 199 $CYAN 1.2 | Out-Null

    Panel 382 282 512 168 $PANEL $GREEN 1.0 0.05 | Out-Null
    Bar 382 282 512 5 $GREEN | Out-Null
    Txt 398 299 480 18 "PROOF AT G6: ONE CLOSED, AUDITABLE LOOP" 9.3 -1 $GREEN "Aptos" 1 1 | Out-Null
    $proof = @(
        @{x=400;t="E3D / Hull`rchange";c=$BLUE},
        @{x=496;t="Affected`rscenarios";c=$ORANGE},
        @{x=592;t="Approved solver`revidence";c=$AMBER},
        @{x=688;t="Human / class`rapproval";c=$VIOLET},
        @{x=784;t="Operations`rfeedback";c=$GREEN}
    )
    foreach ($p in $proof) { Chip $p.x 336 80 46 $p.t $PANEL2 $p.c $WHITE 6.8 }
    for ($i=0; $i -lt 4; $i++) { Arrow ($proof[$i].x+80) 359 ($proof[$i+1].x-3) 359 $CYAN 1.0 | Out-Null }
    Txt 400 402 478 30 "Success means the customer can see, explain, approve and replay every state transition - without replacing certified specialist solvers." 7.8 -1 $LIGHT "Aptos" 1 1 | Out-Null

    Panel 38 472 856 34 "0E2A24" $GREEN 0.9 0.05 | Out-Null
    Txt 54 472 824 34 "WHY NOW  |  The executable reference is ready; the next differentiator is controlled configuration + cross-domain evidence + lifecycle learning." 8.2 -1 $WHITE "Aptos" 1 3 | Out-Null

    # ------------------------------------------------------------------
    # Refine official-signal and evidence language across the deck.
    # ------------------------------------------------------------------
    $sRoadmap = GetSlide "Public roadmap comparison"
    ReplacePrefix $sRoadmap "Industrial AI in CONNECT" "CONNECT industrial intelligence, Unified Engineering, AIM and PI System lifecycle context" | Out-Null
    ReplacePrefix $sRoadmap "Teamcenter X SaaS" "Teamcenter digital twin/thread, lifecycle process automation and AI built for PLM" | Out-Null
    ReplacePrefix $sRoadmap "3DEXPERIENCE cloud" "3DEXPERIENCE / ENOVIA cloud-native collaboration, model-based engineering and virtual twins" | Out-Null
    ReplacePrefix $sRoadmap "Octave asset lifecycle" "Octave asset lifecycle intelligence spanning design, projects, completions and operations" | Out-Null
    ReplacePrefix $sRoadmap "Windchill / Codebeamer" "Windchill + Codebeamer requirements, risk, test, end-to-end traceability and AI assistance" | Out-Null
    ReplacePrefix $sRoadmap "Node Network workflows" "Naval architecture, hydrostatics, stability, rule compliance and calculation workflows" | Out-Null

    $sEvidence = GetSlide "Evidence base and assumption boundaries"
    ReplacePrefix $sEvidence "CONNECT industrial AI" "CONNECT industrial intelligence; Unified Engineering data-centric design; E3D, AIM and PI lifecycle context." | Out-Null
    ReplacePrefix $sEvidence "Teamcenter X SaaS" "Teamcenter digital twin/thread, product-data single source, lifecycle process automation and AI for PLM." | Out-Null
    ReplacePrefix $sEvidence "3DEXPERIENCE cloud" "3DEXPERIENCE / ENOVIA cloud-native model-based collaboration, virtual twins and trusted product data." | Out-Null
    ReplacePrefix $sEvidence "Octave asset lifecycle" "Octave asset lifecycle intelligence across design, projects, completions, operations and maintenance." | Out-Null
    ReplacePrefix $sEvidence "Windchill / Codebeamer" "Windchill + Codebeamer requirements, risk, test, end-to-end traceability and AI-assisted quality." | Out-Null
    ReplacePrefix $sEvidence "Naval architecture and stability" "NAPA naval architecture package: hydrostatics, stability, rule compliance and calculation workflows." | Out-Null
    ReplacePrefix $sEvidence "Key sources:" "Official product pages checked 6 Jun 2026: aveva.com | siemens.com/teamcenter | 3ds.com/3dexperience/enovia | hexagon.com/asset-lifecycle | ptc.com/codebeamer/windchill | napa.fi" | Out-Null

    $sCover = $pres.Slides.Item(1)
    ReplacePrefix $sCover "Current benchmark" "Company-by-company benchmark, win-above strategies, open AVEVA control plane, delivery WBS and KPI gates" | Out-Null

    # Renumber page-number shapes and add any missing number.
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
            Txt 920 9 18 12 ([string]$sl.SlideIndex) 6.5 0 $MUTE "Aptos" 3 1 | Out-Null
        }
    }

    $pres.Save()
    foreach ($sl in $pres.Slides) {
        $sl.Export((Join-Path $reviewDir ("slide-{0:D2}.png" -f $sl.SlideIndex)),"PNG",1280,720)
    }
    Write-Output ("COUNT_AFTER=" + $pres.Slides.Count)
    Write-Output ("REVIEW_DIR=" + $reviewDir)
    Write-Output "FULL_ENHANCEMENT_DONE"
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
