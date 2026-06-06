$ErrorActionPreference = "Stop"

$dir = "D:\이력서\AVEVA - Marine Principal Technical Support & Consultant – PLM SME, Busan\Proposal"
$deck = Join-Path $dir "Future_Industrial_PLM_Meeting_Deck_EN_Rebuilt_v2.pptx"
$backupDir = Join-Path $dir "_backup_20260606"
$reviewDir = "C:\Users\namma\.claude\plm_slide_work\win_above_review_20260606"
New-Item -ItemType Directory -Force -Path $backupDir | Out-Null
New-Item -ItemType Directory -Force -Path $reviewDir | Out-Null
Get-ChildItem -LiteralPath $reviewDir -Filter "slide-*.png" -ErrorAction SilentlyContinue | Remove-Item -Force

$stamp = Get-Date -Format "yyyyMMdd_HHmmss"
$backup = Join-Path $backupDir ("Future_Industrial_PLM_Meeting_Deck_EN_Rebuilt_v2_before_win_above_competitors_" + $stamp + ".pptx")
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
$BGD     = "0B1626"
$PANEL   = "14253B"
$PANEL2  = "102033"
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
        Panel $l $t $w $h $fill $line 0.8 0.30 | Out-Null
        Txt ($l+3) $t ($w-6) $h $text $size 0 $textColor "Aptos" 2 3 | Out-Null
    }

    function BuildHeader($title,$subtitle) {
        $bg = $script:slide.Shapes.AddShape(1,[single]0,[single]0,[single]960,[single]540)
        $bg.Line.Visible = 0
        $bg.Fill.Solid()
        Retry { $bg.Fill.ForeColor.RGB = [int](Rgb $BGD) } | Out-Null
        NoShadow $bg
        Txt 28 16 820 32 $title 21.5 -1 $WHITE "Aptos Display" 1 1 | Out-Null
        Txt 28 51 898 22 $subtitle 9.1 0 $LIGHT "Aptos" 1 1 | Out-Null
        Txt 770 18 158 15 "Future Industrial PLM" 7.5 0 $MUTE "Aptos" 3 1 | Out-Null
        Bar 28 80 900 1 $LINEC | Out-Null
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

    function HasSlide($prefix) {
        foreach ($sl in $pres.Slides) {
            foreach ($sh in $sl.Shapes) {
                try {
                    if ($sh.HasTextFrame -eq -1 -and $sh.TextFrame.HasText -eq -1 -and
                        $sh.TextFrame.TextRange.Text.Trim().StartsWith($prefix)) {
                        return $true
                    }
                } catch {}
            }
        }
        return $false
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

    function SetLayerText($sl,$prefix,$heading,$body) {
        $sh = ReplacePrefix $sl $prefix ($heading + "`r" + $body)
        if ($sh -eq $null) { throw ("Architecture layer not found: " + $prefix) }
        $tr = $sh.TextFrame.TextRange
        $tr.Font.Name = "Aptos"
        $tr.Font.Size = [single]9.2
        $tr.Font.Bold = [int]0
        Retry { $tr.Font.Color.RGB = [int](Rgb $LIGHT) } | Out-Null
        $p1 = $tr.Paragraphs(1,1)
        $p1.Font.Name = "Aptos"
        $p1.Font.Size = [single]10.2
        $p1.Font.Bold = [int]-1
        Retry { $p1.Font.Color.RGB = [int](Rgb $WHITE) } | Out-Null
        $p2 = $tr.Paragraphs(2,1)
        $p2.Font.Name = "Aptos"
        $p2.Font.Size = [single]9.0
        $p2.Font.Bold = [int]0
        Retry { $p2.Font.Color.RGB = [int](Rgb $LIGHT) } | Out-Null
    }

    function AddOuterColumn($x,$accent,$heading,$claim,$chips,$note) {
        Panel $x 104 246 314 $PANEL $accent 0.9 0.05 | Out-Null
        Bar $x 104 246 5 $accent | Out-Null
        Txt ($x+16) 118 214 18 $heading 8.9 -1 $accent "Aptos" 1 1 | Out-Null
        Txt ($x+16) 142 214 42 $claim 12.1 -1 $WHITE "Aptos Display" 1 1 | Out-Null
        $cy = 198
        foreach ($c in $chips) {
            Chip ($x+14) $cy 218 31 $c "102B2B" $accent $LIGHT 7.6
            $cy += 39
        }
        Txt ($x+16) 365 214 43 $note 7.7 0 $LIGHT "Aptos" 1 1 | Out-Null
    }

    function AddGraphPanel($accent,$heading,$claim,$hubLine1,$hubLine2,$nodes,$tag,$footer) {
        Panel 310 104 340 314 $PANEL $accent 1.4 0.05 | Out-Null
        Bar 310 104 340 5 $accent | Out-Null
        Txt 326 118 308 18 $heading 8.9 -1 $accent "Aptos" 2 1 | Out-Null
        Txt 326 141 308 24 $claim 10.9 -1 $WHITE "Aptos Display" 2 1 | Out-Null

        Arrow 480 199 480 225 $LINEC 1.1 | Out-Null
        Arrow 406 238 438 258 $LINEC 1.1 | Out-Null
        Arrow 554 238 522 258 $LINEC 1.1 | Out-Null
        Arrow 406 330 438 294 $LINEC 1.1 | Out-Null
        Arrow 554 330 522 294 $LINEC 1.1 | Out-Null

        Panel 425 225 110 78 "0E2A36" $accent 1.4 0.28 | Out-Null
        Txt 435 239 90 18 $hubLine1 9.0 -1 $accent "Aptos" 2 1 | Out-Null
        Txt 435 258 90 33 $hubLine2 9.0 -1 $WHITE "Aptos" 2 1 | Out-Null

        Panel 422 176 116 30 $PANEL2 $LINEC 0.7 0.22 | Out-Null
        Txt 426 176 108 30 $nodes[0] 7.2 -1 $LIGHT "Aptos" 2 3 | Out-Null
        Panel 330 214 118 34 $PANEL2 $LINEC 0.7 0.22 | Out-Null
        Txt 334 214 110 34 $nodes[1] 7.0 -1 $LIGHT "Aptos" 2 3 | Out-Null
        Panel 512 214 118 34 $PANEL2 $LINEC 0.7 0.22 | Out-Null
        Txt 516 214 110 34 $nodes[2] 7.0 -1 $LIGHT "Aptos" 2 3 | Out-Null
        Panel 330 318 118 34 $PANEL2 $LINEC 0.7 0.22 | Out-Null
        Txt 334 318 110 34 $nodes[3] 7.0 -1 $LIGHT "Aptos" 2 3 | Out-Null
        Panel 512 318 118 34 $PANEL2 $LINEC 0.7 0.22 | Out-Null
        Txt 516 318 110 34 $nodes[4] 7.0 -1 $LIGHT "Aptos" 2 3 | Out-Null

        Chip 425 363 110 24 $tag "2A1431" $MAGENTA $MAGENTA 7.0
        Txt 326 391 308 18 $footer 7.3 0 $MUTE "Aptos" 2 1 | Out-Null
    }

    function AddWinningStrip($accent,$winning,$handoff) {
        Panel 36 432 888 58 "0E2A36" $accent 1.2 0.05 | Out-Null
        Txt 52 443 138 18 "WINNING POSITION" 9.2 -1 $accent "Aptos" 1 1 | Out-Null
        Txt 194 438 714 43 $winning 8.2 -1 $WHITE "Aptos" 1 3 | Out-Null
        Txt 38 503 884 18 $handoff 6.9 0 $MUTE "Aptos" 1 1 | Out-Null
    }

    if (HasSlide "Win above Siemens") { throw "A Win above Siemens slide already exists." }
    if (HasSlide "Win above Dassault") { throw "A Win above Dassault slide already exists." }

    $sAgenda = GetSlide "Meeting structure"
    $sThesis = GetSlide "Meeting thesis"
    $sAbsorb = GetSlide "What to absorb"
    $sNapa = GetSlide "Win above NAPA"
    $sArchitecture = GetSlide "Target architecture concept"
    $sFuture = GetSlide "Future-state comparison"

    ReplacePrefix $sAgenda "Meeting thesis: go beyond competitors" "Meeting thesis + win-above NAPA / Siemens / Dassault patterns" | Out-Null
    ReplacePrefix $sThesis "A vendor-neutral Digital Thread Hub" "An open AVEVA control plane that combines Siemens-grade configuration, Dassault-grade collaboration and NAPA-grade assurance - with customer-owned APIs, evidence and operations feedback." | Out-Null
    ReplacePrefix $sAbsorb "Do not copy a competitor" "The next three slides show how to win above NAPA, Siemens and Dassault - then converge the patterns into one open AVEVA architecture." | Out-Null

    $sSiemens = $pres.Slides.Add([int]($sNapa.SlideIndex + 1),12)
    $script:slide = $sSiemens
    BuildHeader "Win above Siemens - make configuration rigor open and operations-aware" "Teamcenter is strong at PLM governance. Win by preserving that rigor while making configuration customer-owned, API-native and connected to shipyard and operational truth."
    AddOuterColumn 36 $BLUE "SIEMENS STRENGTH TO KEEP" "Preserve enterprise-grade control." @(
        "Item / BOM / change governance",
        "Effectivity + baseline discipline",
        "Enterprise workflow + audit",
        "Multi-domain configuration"
    ) "Keep the rigor that makes complex product change governable."
    AddGraphPanel $CYAN "NEW MOAT: OPEN CONFIGURATION CONTROL PLANE" "A customer-owned graph that can execute the process" "OPEN" "CONFIGURATION`rCONTROL GRAPH" @(
        "Revision / baseline",
        "Hull / block / option / date",
        "ECR / ECO / MCO",
        "M-BOM / production status",
        "Approval / evidence"
    ) "policy-gated APIs" "Event -> policy -> gate -> evidence -> promote"
    AddOuterColumn 678 $GREEN "AVEVA LIFECYCLE CONTEXT" "Connect configuration to real work." @(
        "E3D / Marine design objects",
        "Draw / MTO / ERM production truth",
        "AIM handover + documents",
        "PI / CONNECT operational truth"
    ) "Make every state visible, usable and measurable beyond the PLM specialist."
    Arrow 282 264 306 264 $BLUE 1.6 | Out-Null
    Arrow 674 264 654 264 $GREEN 1.6 | Out-Null
    AddWinningStrip $CYAN "Siemens = governance benchmark. AVEVA-based PLM = open system of configuration action. Competitive moat: customer-owned APIs + live engineering-to-operations context." "ARCHITECTURE HANDOFF  ->  API Contract Layer + Configuration Layer + Integration Layer"

    $sDassault = $pres.Slides.Add([int]($sSiemens.SlideIndex + 1),12)
    $script:slide = $sDassault
    BuildHeader "Win above Dassault - make the virtual twin executable and evidence-driven" "3DEXPERIENCE is strong at immersive collaboration and virtual-twin vision. Win by turning every twin interaction into an API contract, gate result, affected-object trace and operational learning event."
    AddOuterColumn 36 $VIOLET "DASSAULT STRENGTH TO KEEP" "Preserve the shared virtual experience." @(
        "Role-based collaboration spaces",
        "Virtual twin + 3D experience",
        "Engineering / manufacturing view",
        "Visual simulation + validation"
    ) "Keep the context and collaboration that make complex decisions understandable."
    AddGraphPanel $MAGENTA "NEW MOAT: EXECUTABLE EVIDENCE TWIN" "Every interaction becomes a governed next-state event" "EVIDENCE" "TWIN`rCONTROL GRAPH" @(
        "Role / task / decision",
        "Requirement / affected object",
        "AI diagnosis / solver result",
        "Approval / evidence pack",
        "Operations feedback"
    ) "twin action = API event" "Context -> decision -> evidence -> approval -> learning"
    AddOuterColumn 678 $GREEN "AVEVA LIFECYCLE CONTEXT" "Connect the twin to production truth." @(
        "E3D / Marine production truth",
        "AIM document + class context",
        "PI / CONNECT operating feedback",
        "Customer-configured workflows"
    ) "Make the virtual twin drive measurable work, evidence and lifecycle learning."
    Arrow 282 264 306 264 $VIOLET 1.6 | Out-Null
    Arrow 674 264 654 264 $GREEN 1.6 | Out-Null
    AddWinningStrip $MAGENTA "Dassault = collaboration benchmark. AVEVA-based PLM = executable evidence twin. Competitive moat: every virtual-twin decision is measurable, auditable and connected to operations." "ARCHITECTURE HANDOFF  ->  Experience Layer + Assurance & Digital Thread Layer + Integration Layer"

    ReplacePrefix $sArchitecture "A layered, vendor-neutral" "Convergence: Siemens-grade configuration + Dassault-grade collaboration + NAPA-grade assurance inside an open AVEVA control plane" | Out-Null
    SetLayerText $sArchitecture "Experience Layer" "Experience Layer" "Dassault-grade role collaboration, 3D context, assurance margins and executable gate views"
    SetLayerText $sArchitecture "API Contract Layer" "API Contract Layer" "Customer-owned FastAPI/OpenAPI contracts, OIDC authority and solver-adapter contracts"
    SetLayerText $sArchitecture "Assurance + Digital Thread Layer" "Assurance + Digital Thread Layer" "NAPA-grade evidence lineage, assurance graph, rules and evaluate/promote flow"
    SetLayerText $sArchitecture "Configuration Layer" "Configuration Layer" "Siemens-grade BOM, effectivity, baselines and change-impact governance"
    SetLayerText $sArchitecture "Integration Layer" "Integration Layer" "AVEVA E3D/AIM/PI/CONNECT + NAPA/class/CFD/FEM + ERP/MES/PLM adapters"

    ReplacePrefix $sFuture "Proposed AVEVA + Continuous" "Proposed open AVEVA control plane versus NAPA, Siemens / Dassault PLM and other ship design tools" | Out-Null
    ReplacePrefix $sFuture "Classical PLM" "Siemens / Dassault`rPLM platforms" | Out-Null
    ReplacePrefix $sFuture "Distinctive position" "Distinctive position = Siemens-grade configuration + Dassault-grade collaboration + NAPA-grade assurance, connected by customer-owned APIs and AVEVA lifecycle context." | Out-Null

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
    Write-Output "WIN_ABOVE_COMPETITORS_DONE"
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
