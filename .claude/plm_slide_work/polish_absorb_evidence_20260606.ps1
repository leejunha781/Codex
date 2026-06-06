$ErrorActionPreference = "Stop"

$deck = "D:\이력서\AVEVA - Marine Principal Technical Support & Consultant – PLM SME, Busan\Proposal\Future_Industrial_PLM_Meeting_Deck_EN_Rebuilt_v2.pptx"
$reviewDir = "C:\Users\namma\.claude\plm_slide_work\enhanced_full_deck_review_20260606"

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
$GREEN="49C28C";$ORANGE="F0832F";$VIOLET="9B7BE0";$AMBER="E8B23A"
$BGD="0B1626";$PANEL="14253B";$PANEL2="102033";$LINEC="2C4663"
$ppt=$null;$pres=$null

try {
    $ppt=New-Object -ComObject PowerPoint.Application
    $pres=$ppt.Presentations.Open($deck,0,0,0)

    function NoShadow($s){try{$s.Shadow.Visible=0}catch{}}
    function Panel($l,$t,$w,$h,$fill,$line,$lw,$rad){
        $s=$script:slide.Shapes.AddShape(5,[single]$l,[single]$t,[single]$w,[single]$h)
        try{$s.Adjustments.Item(1)=[single]$rad}catch{}
        $s.Fill.Solid();Retry{$s.Fill.ForeColor.RGB=[int](Rgb $fill)}|Out-Null
        if($line){$s.Line.Visible=-1;Retry{$s.Line.ForeColor.RGB=[int](Rgb $line)}|Out-Null;$s.Line.Weight=[single]$lw}else{$s.Line.Visible=0}
        NoShadow $s;return $s
    }
    function Bar($l,$t,$w,$h,$fill){
        $s=$script:slide.Shapes.AddShape(1,[single]$l,[single]$t,[single]$w,[single]$h)
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
    function DelAll($sl){foreach($sh in @($sl.Shapes)){try{$sh.Delete()}catch{}}}
    function BuildHeader($title,$subtitle){
        $bg=$script:slide.Shapes.AddShape(1,[single]0,[single]0,[single]960,[single]540)
        $bg.Line.Visible=0;$bg.Fill.Solid();Retry{$bg.Fill.ForeColor.RGB=[int](Rgb $BGD)}|Out-Null;NoShadow $bg
        Txt 34 18 790 32 $title 21.5 -1 $WHITE "Aptos Display" 1 1|Out-Null
        Txt 34 54 880 20 $subtitle 9.6 0 $LIGHT "Aptos" 1 1|Out-Null
        Txt 800 24 130 15 "Future Industrial PLM" 7.6 0 $MUTE "Aptos" 3 1|Out-Null
        Bar 34 82 892 1 $LINEC|Out-Null
    }

    # Explicitly retain every key differentiator from the previous absorption page.
    $sAbsorb=$pres.Slides.Item(9)
    $diff=$sAbsorb.Shapes.Item("TextBox 114")
    $diff.TextFrame.TextRange.Text="E3D / Marine + Unified Engineering + AIM + PI + CONNECT`rCustomer-owned OpenAPI + configuration / assurance graph`rFederated solver adapters + physics-guarded AI + signed evidence / operations learning"
    $diff.TextFrame.TextRange.Font.Name="Aptos"
    $diff.TextFrame.TextRange.Font.Size=[single]6.7
    $diff.TextFrame.TextRange.Font.Bold=[int]-1

    # Use the available whitespace in the pilot cards for better presentation readability.
    foreach($name in @("TextBox 69","TextBox 73","TextBox 77")){
        $sh=$pres.Slides.Item(16).Shapes.Item($name)
        $sh.TextFrame.TextRange.Font.Size=[single]8.7
    }

    # Rebuild the evidence page to eliminate any latent template/background artifacts.
    $sEvidence=$pres.Slides.Item(29)
    DelAll $sEvidence
    $script:slide=$sEvidence
    BuildHeader "Evidence base and assumption boundaries" "Official product signals support the direction; proposed AVEVA control-plane capabilities still require pilot validation"

    $evidence=@(
        @{x=38;y=112;c=$CYAN;v="AVEVA";s="PUBLIC SIGNAL`rData-centric engineering + industrial intelligence";u="DESIGN USE`rEngineering-to-operations lifecycle context"},
        @{x=346;y=112;c=$BLUE;v="SIEMENS";s="PUBLIC SIGNAL`rDigital thread / twin + PLM process automation";u="DESIGN USE`rConfiguration rigor and lifecycle governance"},
        @{x=654;y=112;c=$VIOLET;v="DASSAULT";s="PUBLIC SIGNAL`rModel-based collaboration + virtual twins";u="DESIGN USE`rRole context and executable twin decisions"},
        @{x=38;y=270;c=$GREEN;v="HEXAGON";s="PUBLIC SIGNAL`rAsset lifecycle from design through maintenance";u="DESIGN USE`rCompletions, handover and operations context"},
        @{x=346;y=270;c=$ORANGE;v="PTC";s="PUBLIC SIGNAL`rRequirements, risk, test traceability + AI";u="DESIGN USE`rVerification gates and affected-object trace"},
        @{x=654;y=270;c=$AMBER;v="NAPA";s="PUBLIC SIGNAL`rNaval calculations, rules and workflows";u="DESIGN USE`rApproved solver-of-record adapter"}
    )
    foreach($e in $evidence){
        Panel $e.x $e.y 270 140 $PANEL $e.c 0.9 0.05|Out-Null
        Bar $e.x $e.y 270 5 $e.c|Out-Null
        Txt ($e.x+14) ($e.y+14) 242 18 $e.v 9.2 -1 $e.c "Aptos" 1 1|Out-Null
        Txt ($e.x+14) ($e.y+42) 242 38 $e.s 7.1 0 $LIGHT "Aptos" 1 1|Out-Null
        Bar ($e.x+14) ($e.y+88) 242 1 $LINEC|Out-Null
        Txt ($e.x+14) ($e.y+98) 242 34 $e.u 7.1 -1 $WHITE "Aptos" 1 1|Out-Null
    }
    Panel 38 432 886 54 "2A2510" $AMBER 0.9 0.05|Out-Null
    Txt 54 432 132 54 "ASSUMPTION`rBOUNDARY" 8.5 -1 $AMBER "Aptos" 1 3|Out-Null
    Txt 194 434 714 50 "Public signals are not product commitments. The assurance graph, solver adapters, AI diagnosis and closed-loop calibration are proposed AVEVA-based PLM extensions; validate scope, authority and evidence with HD Hyundai, vendors and class authorities." 7.6 -1 $WHITE "Aptos" 1 3|Out-Null
    Txt 40 505 884 13 "Official product pages checked 6 Jun 2026: aveva.com | siemens.com/teamcenter | 3ds.com/3dexperience/enovia | hexagon.com/asset-lifecycle | ptc.com/codebeamer/windchill | napa.fi" 5.9 0 $MUTE "Aptos" 1 1|Out-Null

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
    Write-Output "POLISH_ABSORB_EVIDENCE_DONE"
}
finally {
    if($pres){try{$pres.Close()}catch{};try{[Runtime.InteropServices.Marshal]::ReleaseComObject($pres)|Out-Null}catch{}}
    if($ppt){try{if($ppt.Presentations.Count-eq 0){$ppt.Quit()}}catch{};try{[Runtime.InteropServices.Marshal]::ReleaseComObject($ppt)|Out-Null}catch{}}
    [GC]::Collect();[GC]::WaitForPendingFinalizers()
}
