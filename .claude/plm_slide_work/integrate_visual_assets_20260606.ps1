$ErrorActionPreference = "Stop"

$dir = "D:\이력서\AVEVA - Marine Principal Technical Support & Consultant – PLM SME, Busan\Proposal"
$deck = Join-Path $dir "Future_Industrial_PLM_Meeting_Deck_EN_Rebuilt_v2.pptx"
$assetDir = Join-Path $dir "plm_integrated_english_package\04_Presentation Image"
$generatedDir = Join-Path $assetDir "Generated_20260606"
$controlPlaneImg = Join-Path $generatedDir "08_open_aveva_control_plane_generated.png"
$assuranceLoopImg = Join-Path $generatedDir "16_continuous_naval_assurance_loop_generated.png"
$boardroomImg = Join-Path $assetDir "02_a_wide_cinematic_corporate_boardroom_command_cent.png"
$closingImg = Join-Path $assetDir "04_a_wide_cinematic_industrial_technology_visualizat.png"
$backupDir = Join-Path $dir "_backup_20260606"
$reviewDir = "C:\Users\namma\.claude\plm_slide_work\visual_integration_review_20260606"

foreach ($p in @($controlPlaneImg,$assuranceLoopImg,$boardroomImg,$closingImg)) {
    if (-not (Test-Path -LiteralPath $p)) { throw ("Missing image: " + $p) }
}
New-Item -ItemType Directory -Force -Path $backupDir | Out-Null
New-Item -ItemType Directory -Force -Path $reviewDir | Out-Null
Get-ChildItem -LiteralPath $reviewDir -Filter "slide-*.png" -ErrorAction SilentlyContinue | Remove-Item -Force

$stamp = Get-Date -Format "yyyyMMdd_HHmmss"
$backup = Join-Path $backupDir ("Future_Industrial_PLM_Meeting_Deck_EN_Rebuilt_v2_before_visual_integration_" + $stamp + ".pptx")
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
    for ($a=0; $a -lt 8; $a++) { try { return (& $sb) } catch { Start-Sleep -Milliseconds 120 } }
    return (& $sb)
}

$WHITE="FFFFFF";$LIGHT="C7D6E5";$MUTE="93A8C0";$CYAN="34C0EE";$BLUE="4F8BE8"
$GREEN="49C28C";$ORANGE="F0832F";$VIOLET="9B7BE0";$MAGENTA="E653B5";$AMBER="E8B23A"
$BGD="0B1626";$PANEL="14253B";$PANEL2="102033";$LINEC="2C4663"
$ppt=$null;$pres=$null

try {
    $ppt=New-Object -ComObject PowerPoint.Application
    $pres=$ppt.Presentations.Open($deck,0,0,0)

    function NoShadow($s){try{$s.Shadow.Visible=0}catch{}}
    function Panel($l,$t,$w,$h,$fill,$line,$lw,$rad,$transparency=0){
        $s=$script:slide.Shapes.AddShape(5,[single]$l,[single]$t,[single]$w,[single]$h)
        try{$s.Adjustments.Item(1)=[single]$rad}catch{}
        $s.Fill.Solid();Retry{$s.Fill.ForeColor.RGB=[int](Rgb $fill)}|Out-Null
        $s.Fill.Transparency=[single]$transparency
        if($line){$s.Line.Visible=-1;Retry{$s.Line.ForeColor.RGB=[int](Rgb $line)}|Out-Null;$s.Line.Weight=[single]$lw}else{$s.Line.Visible=0}
        NoShadow $s;return $s
    }
    function Rect($l,$t,$w,$h,$fill,$transparency){
        $s=$script:slide.Shapes.AddShape(1,[single]$l,[single]$t,[single]$w,[single]$h)
        $s.Fill.Solid();Retry{$s.Fill.ForeColor.RGB=[int](Rgb $fill)}|Out-Null;$s.Fill.Transparency=[single]$transparency
        $s.Line.Visible=0;NoShadow $s;return $s
    }
    function Bar($l,$t,$w,$h,$fill){
        $s=$script:slide.Shapes.AddShape(1,[single]$l,[single]$t,[single]$w,[single]$h)
        $s.Fill.Solid();Retry{$s.Fill.ForeColor.RGB=[int](Rgb $fill)}|Out-Null;$s.Line.Visible=0;NoShadow $s;return $s
    }
    function Oval($l,$t,$w,$h,$fill){
        $s=$script:slide.Shapes.AddShape(9,[single]$l,[single]$t,[single]$w,[single]$h)
        $s.Fill.Solid();Retry{$s.Fill.ForeColor.RGB=[int](Rgb $fill)}|Out-Null;$s.Line.Visible=0;NoShadow $s;return $s
    }
    function Txt($l,$t,$w,$h,$text,$size,$bold,$color,$font,$align,$anchor){
        $tb=$script:slide.Shapes.AddTextbox(1,[single]$l,[single]$t,[single]$w,[single]$h)
        $tb.Fill.Visible=0;$tb.Line.Visible=0;NoShadow $tb
        $tf=$tb.TextFrame;$tf.WordWrap=-1;$tf.AutoSize=0
        $tf.MarginLeft=[single]2;$tf.MarginRight=[single]2;$tf.MarginTop=[single]1;$tf.MarginBottom=[single]1;$tf.VerticalAnchor=[int]$anchor
        $tr=$tf.TextRange;$tr.Text=[string]$text;$tr.Font.Size=[single]$size;$tr.Font.Bold=[int]$bold;$tr.Font.Name=[string]$font
        Retry{$tr.Font.Color.RGB=[int](Rgb $color)}|Out-Null;$tr.ParagraphFormat.Alignment=[int]$align
        try{$tr.ParagraphFormat.SpaceBefore=[single]0;$tr.ParagraphFormat.SpaceAfter=[single]0}catch{}
        return $tb
    }
    function Chip($l,$t,$w,$h,$text,$fill,$line,$textColor,$size){
        Panel $l $t $w $h $fill $line 0.8 0.25 0.05|Out-Null
        Txt ($l+3) $t ($w-6) $h $text $size 0 $textColor "Aptos" 2 3|Out-Null
    }
    function AddPicture($path,$l,$t,$w,$h){
        $pic=$script:slide.Shapes.AddPicture($path,0,-1,[single]$l,[single]$t,[single]$w,[single]$h)
        NoShadow $pic
        return $pic
    }
    function DelAll($sl){foreach($sh in @($sl.Shapes)){try{$sh.Delete()}catch{}}}
    function BuildHeader($title,$subtitle){
        Txt 34 18 790 32 $title 21.5 -1 $WHITE "Aptos Display" 1 1|Out-Null
        Txt 34 54 880 20 $subtitle 9.6 0 $LIGHT "Aptos" 1 1|Out-Null
        Txt 800 24 130 15 "Future Industrial PLM" 7.6 0 $MUTE "Aptos" 3 1|Out-Null
        Bar 34 82 892 1 $LINEC|Out-Null
    }

    # ------------------------------------------------------------------
    # Slide 9: use a generated visual as the proof object for convergence.
    # ------------------------------------------------------------------
    $sAbsorb=$pres.Slides.Item(9)
    foreach($sh in @($sAbsorb.Shapes)){
        try{
            if($sh.Left -ge 556 -and $sh.Top -ge 110 -and $sh.Top -lt 445){$sh.Delete()}
        }catch{}
    }
    $script:slide=$sAbsorb
    AddPicture $controlPlaneImg 558 116 366 206|Out-Null
    Rect 558 116 366 206 $BGD 0.36|Out-Null
    Panel 558 116 366 206 $BGD $CYAN 1.2 0.04 1|Out-Null
    Bar 558 116 366 6 $CYAN|Out-Null
    Txt 576 132 328 20 "NEW AVEVA-BASED PLM" 11 -1 $CYAN "Aptos Display" 1 1|Out-Null
    Txt 576 158 328 34 "Five specialist strengths -> one open decision control plane" 9.3 -1 $WHITE "Aptos" 1 1|Out-Null
    Panel 558 330 366 104 "0E2436" $CYAN 1.0 0.04 0.05|Out-Null
    Txt 576 342 328 18 "AVEVA DIFFERENTIATION" 8.8 -1 $CYAN "Aptos" 1 1|Out-Null
    $moat=@(
        "E3D / Marine + Unified Engineering + AIM + PI + CONNECT",
        "Customer-owned OpenAPI + configuration / assurance graph",
        "Federated solvers + physics-guarded AI + evidence + operations learning"
    )
    $my=366
    foreach($m in $moat){
        Oval 578 ($my+4) 7 7 $GREEN|Out-Null
        Txt 592 $my 310 18 $m 6.8 -1 $LIGHT "Aptos" 1 1|Out-Null
        $my+=22
    }

    # ------------------------------------------------------------------
    # Slide 16: combine the executable workflow with the generated loop.
    # ------------------------------------------------------------------
    $sPilot=$pres.Slides.Item(16)
    foreach($sh in @($sPilot.Shapes)){
        try{
            if($sh.Top -ge 238 -and $sh.Top -lt 510){$sh.Delete()}
        }catch{}
    }
    $script:slide=$sPilot
    AddPicture $assuranceLoopImg 38 245 520 293|Out-Null
    Rect 38 245 520 293 $BGD 0.28|Out-Null
    Panel 38 245 520 293 $BGD $CYAN 1.0 0.04 1|Out-Null
    Txt 54 258 474 18 "VISUAL PROOF: ONE CLOSED ASSURANCE LOOP" 9.2 -1 $CYAN "Aptos" 1 1|Out-Null
    Txt 54 476 474 36 "Configuration -> affected scenarios -> approved solver evidence -> human / class approval -> operations feedback" 7.4 -1 $WHITE "Aptos" 1 1|Out-Null
    Txt 54 514 474 13 "AI diagnoses and holds gates; accountable humans retain release authority." 6.4 0 $MUTE "Aptos" 1 1|Out-Null

    Panel 580 245 342 82 $PANEL $BLUE 0.9 0.04 0.03|Out-Null
    Bar 580 245 342 5 $BLUE|Out-Null
    Txt 596 258 310 17 "BOUNDED SCENARIO" 8.7 -1 $BLUE "Aptos" 1 1|Out-Null
    Txt 596 282 310 34 "One hull / block · selected loading cases · one approved solver adapter · seeded change and failure cases" 7.2 0 $LIGHT "Aptos" 1 1|Out-Null

    Panel 580 337 342 82 $PANEL $GREEN 0.9 0.04 0.03|Out-Null
    Bar 580 337 342 5 $GREEN|Out-Null
    Txt 596 350 310 17 "DELIVERABLE + RELEASE GOVERNANCE" 8.7 -1 $GREEN "Aptos" 1 1|Out-Null
    Txt 596 374 310 34 "Configuration + assurance graph · signed evidence pack · AI proposes / holds · naval architect and class authority release" 7.2 0 $LIGHT "Aptos" 1 1|Out-Null

    Panel 580 429 342 109 $PANEL $MAGENTA 0.9 0.04 0.03|Out-Null
    Bar 580 429 342 5 $MAGENTA|Out-Null
    Txt 596 442 310 17 "MINIMUM ACCEPTANCE" 8.7 -1 $MAGENTA "Aptos" 1 1|Out-Null
    Chip 596 468 145 27 "Configuration >=99%" $PANEL2 $GREEN $WHITE 6.7
    Chip 757 468 145 27 "Impact recall >=95%" $PANEL2 $ORANGE $WHITE 6.7
    Chip 596 502 145 27 "Provenance 100%" $PANEL2 $MAGENTA $WHITE 6.7
    Chip 757 502 145 27 "Approval <=5 s" $PANEL2 $BLUE $WHITE 6.7

    # ------------------------------------------------------------------
    # Slide 27: boardroom image becomes the operating-model context.
    # ------------------------------------------------------------------
    $sReview=$pres.Slides.Item(27)
    DelAll $sReview
    $script:slide=$sReview
    AddPicture $boardroomImg 0 0 960 540|Out-Null
    Rect 0 0 960 540 $BGD 0.38|Out-Null
    Rect 0 0 960 104 $BGD 0.08|Out-Null
    BuildHeader "Review ownership and operating model" "A cross-functional review board sees the same KPI, risk, evidence and release state before any capability is promoted"

    Txt 42 112 260 18 "ACCOUNTABLE REVIEW BOARD" 9 -1 $CYAN "Aptos" 1 1|Out-Null
    $roles=@(
        @{x=42;c=$CYAN;t="PRODUCT PLANNING";d="Roadmap · MVP scope · user outcomes · KPI thresholds"},
        @{x=270;c=$AMBER;t="NAVAL ARCHITECTURE + CLASS";d="Criteria · solver validation · evidence adequacy · release"},
        @{x=498;c=$GREEN;t="SOLUTION ARCHITECTURE";d="Data model · assurance graph · OpenAPI · adapters · NFRs"},
        @{x=726;c=$VIOLET;t="DEVELOPMENT TEAM";d="Implementation · automation · security · performance · defects"}
    )
    foreach($r in $roles){
        Panel $r.x 306 198 102 $PANEL $r.c 0.9 0.04 0.10|Out-Null
        Bar $r.x 306 198 5 $r.c|Out-Null
        Txt ($r.x+12) 322 174 18 $r.t 7.7 -1 $r.c "Aptos" 1 1|Out-Null
        Txt ($r.x+12) 351 174 43 $r.d 7.0 -1 $WHITE "Aptos" 1 1|Out-Null
    }
    Panel 42 428 882 74 "0E2436" $CYAN 1.0 0.04 0.04|Out-Null
    Txt 58 440 190 18 "WEEKLY STEERING + REVIEW" 9.2 -1 $CYAN "Aptos" 1 1|Out-Null
    Txt 256 438 650 42 "1  KPI + evidence dashboard     2  Failing configuration / assurance gates     3  Class and naval-architecture issues`r4  Scope-change decisions     5  Promote completed capabilities + update risk burndown" 7.3 -1 $WHITE "Aptos" 1 1|Out-Null
    Txt 920 9 18 12 "27" 6.5 0 $MUTE "Aptos" 3 1|Out-Null

    # ------------------------------------------------------------------
    # Slide 30: cinematic closing background while retaining decision proof.
    # ------------------------------------------------------------------
    $sDecision=$pres.Slides.Item(30)
    $script:slide=$sDecision
    $overlay=Rect 0 0 960 540 $BGD 0.34
    $overlay.ZOrder(1)
    $pic=AddPicture $closingImg 0 0 960 540
    $pic.ZOrder(1)
    # Increase opacity of the existing decision surfaces over the image.
    foreach($sh in $sDecision.Shapes){
        try{
            if($sh.Type -eq 1 -and $sh.Width -gt 75 -and $sh.Height -gt 30 -and $sh.Top -gt 90){
                $sh.Fill.Transparency=[single]0.05
            }
        }catch{}
    }

    # Ensure all page numbers remain correct.
    foreach($sl in $pres.Slides){
        $updated=$false
        foreach($sh in $sl.Shapes){
            try{
                if($sh.HasTextFrame -eq -1 -and $sh.TextFrame.HasText -eq -1 -and $sh.Left -gt 890 -and $sh.Top -lt 55 -and $sh.Width -lt 60 -and $sh.Height -lt 35){
                    $sh.TextFrame.TextRange.Text=[string]$sl.SlideIndex;$updated=$true
                }
            }catch{}
        }
        if(-not $updated){$script:slide=$sl;Txt 920 9 18 12 ([string]$sl.SlideIndex) 6.5 0 $MUTE "Aptos" 3 1|Out-Null}
    }

    $pres.Save()
    foreach($sl in $pres.Slides){$sl.Export((Join-Path $reviewDir ("slide-{0:D2}.png" -f $sl.SlideIndex)),"PNG",1280,720)}
    Write-Output ("COUNT_AFTER=" + $pres.Slides.Count)
    Write-Output ("REVIEW_DIR=" + $reviewDir)
    Write-Output "VISUAL_INTEGRATION_DONE"
}
finally {
    if($pres){try{$pres.Close()}catch{};try{[Runtime.InteropServices.Marshal]::ReleaseComObject($pres)|Out-Null}catch{}}
    if($ppt){try{if($ppt.Presentations.Count-eq 0){$ppt.Quit()}}catch{};try{[Runtime.InteropServices.Marshal]::ReleaseComObject($ppt)|Out-Null}catch{}}
    [GC]::Collect();[GC]::WaitForPendingFinalizers()
}
