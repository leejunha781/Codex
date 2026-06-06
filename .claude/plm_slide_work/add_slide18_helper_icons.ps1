$ErrorActionPreference = "Stop"

$dir = "D:\이력서\AVEVA - Marine Principal Technical Support & Consultant – PLM SME, Busan\Proposal"
$deck = Join-Path $dir "Future_Industrial_PLM_Meeting_Deck_EN_Rebuilt_v2_VisualReviewed.pptx"
$stamp = Get-Date -Format "yyyyMMdd_HHmmss"
$backupDir = Join-Path $dir ("_backup_" + $stamp)
$backup = Join-Path $backupDir ("Future_Industrial_PLM_Meeting_Deck_EN_Rebuilt_v2_VisualReviewed_BEFORE_SLIDE18_ICONS_" + $stamp + ".pptx")
$render = "C:\Users\namma\.claude\plm_slide_work\visualreview_current\slide18_icons.png"

New-Item -ItemType Directory -Force -Path $backupDir | Out-Null
Copy-Item -LiteralPath $deck -Destination $backup -Force

function Rgb([string]$hex) {
    $hex=$hex.TrimStart('#')
    return [int](
        [Convert]::ToInt32($hex.Substring(0,2),16) +
        [Convert]::ToInt32($hex.Substring(2,2),16)*256 +
        [Convert]::ToInt32($hex.Substring(4,2),16)*65536
    )
}
function Set-Line($s,$color,$weight) {
    $s.Fill.Visible=0
    $s.Line.Visible=-1
    $s.Line.ForeColor.RGB=[int](Rgb $color)
    $s.Line.Weight=[single]$weight
}
function Solid($s,$fill,$line,$weight) {
    $s.Fill.Solid(); $s.Fill.ForeColor.RGB=[int](Rgb $fill)
    $s.Line.Visible=-1; $s.Line.ForeColor.RGB=[int](Rgb $line); $s.Line.Weight=[single]$weight
}
function Add-Line($sl,$x1,$y1,$x2,$y2,$color,$weight) {
    $l=$sl.Shapes.AddLine([single]$x1,[single]$y1,[single]$x2,[single]$y2)
    $l.Line.ForeColor.RGB=[int](Rgb $color); $l.Line.Weight=[single]$weight
    return $l
}
function Add-Plate($sl,$x,$y,$accent,$num) {
    $p=$sl.Shapes.AddShape(5,[single]$x,[single]$y,[single]48,[single]22)
    $p.Name=("HelperIcon{0}_Plate" -f $num)
    $p.Fill.Solid(); $p.Fill.ForeColor.RGB=[int](Rgb "0B1A2C"); $p.Fill.Transparency=[single]0.05
    $p.Line.Visible=-1; $p.Line.ForeColor.RGB=[int](Rgb $accent); $p.Line.Transparency=[single]0.15; $p.Line.Weight=[single]0.75
    return $p
}
function Name-It($s,$num,$part){ $s.Name=("HelperIcon{0}_{1}" -f $num,$part) }

$ppt=$null;$pres=$null
try {
    $ppt=New-Object -ComObject PowerPoint.Application
    $pres=$ppt.Presentations.Open($deck,$false,$false,$false)
    $sl=$pres.Slides.Item(18)

    # Remove prior helper icons if the script is re-run.
    for($i=$sl.Shapes.Count;$i -ge 1;$i--){
        if($sl.Shapes.Item($i).Name -like "HelperIcon*"){ $sl.Shapes.Item($i).Delete() }
    }

    # Each icon occupies the lower-right space between the description and status chip.
    $cards=@(
        @{n=1;x=222;y=172;c="31C5F4";kind="cube"},
        @{n=2;x=542;y=172;c="31C5F4";kind="drawing"},
        @{n=3;x=862;y=172;c="31C5F4";kind="bom"},
        @{n=4;x=862;y=316;c="4F8BE8";kind="approve"},
        @{n=5;x=542;y=316;c="49C28C";kind="scope"},
        @{n=6;x=222;y=316;c="49C28C";kind="baseline"},
        @{n=7;x=222;y=458;c="F0832F";kind="change"},
        @{n=8;x=542;y=458;c="F0832F";kind="factory"},
        @{n=9;x=862;y=458;c="9B7BE0";kind="impact"}
    )

    foreach($d in $cards){
        $x=[double]$d.x; $y=[double]$d.y; $c=[string]$d.c; $n=[int]$d.n
        Add-Plate $sl $x $y $c $n | Out-Null

        switch($d.kind){
            "cube" {
                $a=$sl.Shapes.AddShape(1,$x+12,$y+6,15,11); Set-Line $a $c 1.1; Name-It $a $n "CubeFront"
                $b=Add-Line $sl ($x+12) ($y+6) ($x+20) ($y+2) $c 1.1; Name-It $b $n "CubeTop1"
                $b=Add-Line $sl ($x+27) ($y+6) ($x+35) ($y+2) $c 1.1; Name-It $b $n "CubeTop2"
                $b=Add-Line $sl ($x+20) ($y+2) ($x+35) ($y+2) $c 1.1; Name-It $b $n "CubeTop3"
                $b=Add-Line $sl ($x+27) ($y+17) ($x+35) ($y+13) $c 1.1; Name-It $b $n "CubeSide1"
                $b=Add-Line $sl ($x+35) ($y+2) ($x+35) ($y+13) $c 1.1; Name-It $b $n "CubeSide2"
            }
            "drawing" {
                $a=$sl.Shapes.AddShape(1,$x+14,$y+3,20,16); Set-Line $a $c 1.0; Name-It $a $n "Sheet"
                foreach($yy in @(8,12,16)){ $l=Add-Line $sl ($x+18) ($y+$yy) ($x+30) ($y+$yy) $c 0.8; Name-It $l $n ("Dim"+$yy) }
                $l=Add-Line $sl ($x+31) ($y+3) ($x+34) ($y+6) $c 0.9; Name-It $l $n "Corner1"
            }
            "bom" {
                foreach($i in 0..2){
                    $o=$sl.Shapes.AddShape(9,$x+8,$y+4+($i*6),4,4); Solid $o $c $c 0.5; Name-It $o $n ("Node"+$i)
                    $r=$sl.Shapes.AddShape(1,$x+17,$y+4+($i*6),21,4); Solid $r "17304B" $c 0.6; Name-It $r $n ("Row"+$i)
                }
                $l=Add-Line $sl ($x+10) ($y+6) ($x+10) ($y+18) $c 0.7; Name-It $l $n "Tree"
            }
            "approve" {
                $o=$sl.Shapes.AddShape(9,$x+13,$y+3,19,17); Set-Line $o $c 1.1; Name-It $o $n "Stamp"
                $l=Add-Line $sl ($x+17) ($y+12) ($x+21) ($y+16) $c 1.5; Name-It $l $n "Check1"
                $l=Add-Line $sl ($x+21) ($y+16) ($x+29) ($y+7) $c 1.5; Name-It $l $n "Check2"
            }
            "scope" {
                foreach($i in 0..3){
                    $r=$sl.Shapes.AddShape(1,$x+8+($i*8),$y+7,6,8)
                    if($i -eq 1){ Solid $r $c $c 0.6 } else { Set-Line $r $c 0.8 }
                    Name-It $r $n ("Segment"+$i)
                }
                $l=Add-Line $sl ($x+6) ($y+18) ($x+40) ($y+18) $c 0.7; Name-It $l $n "Timeline"
            }
            "baseline" {
                $r=$sl.Shapes.AddShape(1,$x+9,$y+10,28,9); Set-Line $r $c 1.0; Name-It $r $n "Snapshot"
                $a=$sl.Shapes.AddShape(9,$x+18,$y+2,10,10); Set-Line $a $c 1.1; Name-It $a $n "LockLoop"
                $b=$sl.Shapes.AddShape(1,$x+17,$y+7,12,10); Solid $b "17304B" $c 1.0; Name-It $b $n "LockBody"
            }
            "change" {
                $l=Add-Line $sl ($x+8) ($y+7) ($x+37) ($y+7) $c 1.2; $l.Line.EndArrowheadStyle=3; Name-It $l $n "Forward"
                $l=Add-Line $sl ($x+37) ($y+16) ($x+8) ($y+16) $c 1.2; $l.Line.EndArrowheadStyle=3; Name-It $l $n "Back"
                $o=$sl.Shapes.AddShape(9,$x+20,$y+8,7,7); Solid $o $c $c 0.6; Name-It $o $n "ChangeNode"
            }
            "factory" {
                $r=$sl.Shapes.AddShape(1,$x+9,$y+10,27,9); Set-Line $r $c 1.0; Name-It $r $n "Base"
                $l=Add-Line $sl ($x+9) ($y+10) ($x+16) ($y+5) $c 1.0; Name-It $l $n "Roof1"
                $l=Add-Line $sl ($x+16) ($y+5) ($x+23) ($y+10) $c 1.0; Name-It $l $n "Roof2"
                $l=Add-Line $sl ($x+23) ($y+10) ($x+29) ($y+5) $c 1.0; Name-It $l $n "Roof3"
                $l=Add-Line $sl ($x+29) ($y+5) ($x+36) ($y+10) $c 1.0; Name-It $l $n "Roof4"
                foreach($xx in @(14,22,30)){ $o=$sl.Shapes.AddShape(9,$x+$xx,$y+13,3,3); Solid $o $c $c 0.4; Name-It $o $n ("Part"+$xx) }
            }
            "impact" {
                $center=$sl.Shapes.AddShape(9,$x+20,$y+8,8,8); Solid $center $c $c 0.6; Name-It $center $n "Center"
                $pts=@(@(10,5),@(36,5),@(10,17),@(36,17))
                $k=0
                foreach($pt in $pts){
                    $o=$sl.Shapes.AddShape(9,$x+$pt[0],$y+$pt[1],5,5); Set-Line $o $c 0.8; Name-It $o $n ("Affected"+$k)
                    $l=Add-Line $sl ($x+24) ($y+12) ($x+$pt[0]+2.5) ($y+$pt[1]+2.5) $c 0.7; Name-It $l $n ("Link"+$k)
                    $k++
                }
                $l=Add-Line $sl ($x+21) ($y+12) ($x+23) ($y+14) "FFFFFF" 1.1; Name-It $l $n "Check1"
                $l=Add-Line $sl ($x+23) ($y+14) ($x+27) ($y+9) "FFFFFF" 1.1; Name-It $l $n "Check2"
            }
        }
    }

    $pres.Save()
    $sl.Export($render,"PNG",1600,900)
}
finally {
    if($pres){try{$pres.Close()}catch{}}
    if($ppt){try{if($ppt.Presentations.Count-eq 0){$ppt.Quit()}}catch{}}
    [GC]::Collect();[GC]::WaitForPendingFinalizers()
}
Write-Output ("BACKUP="+$backup)
Write-Output ("RENDER="+$render)
