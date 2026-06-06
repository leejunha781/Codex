$ErrorActionPreference = "Stop"

$dir = "D:\이력서\AVEVA - Marine Principal Technical Support & Consultant – PLM SME, Busan\Proposal"
$deck = Join-Path $dir "Future_Industrial_PLM_Meeting_Deck_EN_Rebuilt_v2_VisualReviewed.pptx"
$img = Join-Path $dir "plm_integrated_english_package\04_Presentation Image\a_wide_high_resolution_futuristic_infographic_d.png"
$stamp = Get-Date -Format "yyyyMMdd_HHmmss"
$backupDir = Join-Path $dir ("_backup_" + $stamp)
$backup = Join-Path $backupDir ("Future_Industrial_PLM_Meeting_Deck_EN_Rebuilt_v2_VisualReviewed_BEFORE_REFERENCE_ICONS_" + $stamp + ".pptx")
$render = "C:\Users\namma\.claude\plm_slide_work\visualreview_current\slide18_reference_icons.png"

New-Item -ItemType Directory -Force -Path $backupDir | Out-Null
Copy-Item -LiteralPath $deck -Destination $backup -Force
if(-not(Test-Path -LiteralPath $img)){throw "Reference image missing: $img"}

function Rgb([string]$hex) {
    $hex=$hex.TrimStart('#')
    [int]([Convert]::ToInt32($hex.Substring(0,2),16)+[Convert]::ToInt32($hex.Substring(2,2),16)*256+[Convert]::ToInt32($hex.Substring(4,2),16)*65536)
}
function LineStyle($s,$c,$w){$s.Line.Visible=-1;$s.Line.ForeColor.RGB=[int](Rgb $c);$s.Line.Weight=[single]$w}
function NoFillLine($s,$c,$w){$s.Fill.Visible=0;LineStyle $s $c $w}
function FillLine($s,$fill,$c,$w){$s.Fill.Solid();$s.Fill.ForeColor.RGB=[int](Rgb $fill);LineStyle $s $c $w}
function L($sl,$x1,$y1,$x2,$y2,$c,$w,$name){
    $s=$sl.Shapes.AddLine([single]$x1,[single]$y1,[single]$x2,[single]$y2);LineStyle $s $c $w;$s.Name=$name;return $s
}
function N($s,$num,$part){$s.Name=("RefIcon{0}_{1}"-f$num,$part)}
function Plate($sl,$x,$y,$c,$num){
    $p=$sl.Shapes.AddShape(5,[single]$x,[single]$y,[single]48,[single]22)
    $p.Name=("RefIcon{0}_Plate"-f$num);$p.Fill.Solid();$p.Fill.ForeColor.RGB=[int](Rgb "071728");$p.Fill.Transparency=[single]0.03
    LineStyle $p $c 0.85;return $p
}

$ppt=$null;$pres=$null
try{
    $ppt=New-Object -ComObject PowerPoint.Application
    $pres=$ppt.Presentations.Open($deck,$false,$false,$false)
    $sl=$pres.Slides.Item(18)

    # Reference image as a subtle background; content panel remains dominant.
    $sl.FollowMasterBackground=$false
    $sl.Background.Fill.UserPicture($img)
    for($i=$sl.Shapes.Count;$i-ge1;$i--){
        if($sl.Shapes.Item($i).Name -eq "Slide18BgScrim"){$sl.Shapes.Item($i).Delete()}
    }
    $bgScrim=$sl.Shapes.AddShape(1,0,0,960,540)
    $bgScrim.Name="Slide18BgScrim";$bgScrim.Fill.Solid();$bgScrim.Fill.ForeColor.RGB=[int](Rgb "03101F");$bgScrim.Fill.Transparency=[single]0.22
    $bgScrim.Line.Visible=0;$bgScrim.ZOrder(1)
    $panel=$sl.Shapes.Item("Rectangle 6")
    $panel.Fill.Solid();$panel.Fill.ForeColor.RGB=[int](Rgb "071728");$panel.Fill.Transparency=[single]0.02

    for($i=$sl.Shapes.Count;$i-ge1;$i--){
        $nm=$sl.Shapes.Item($i).Name
        if($nm-like"HelperIcon*" -or $nm-like"RefIcon*"){$sl.Shapes.Item($i).Delete()}
    }

    $icons=@(
        @{n=1;x=222;y=172;c="31C5F4";k="ship"},
        @{n=2;x=542;y=172;c="31C5F4";k="drawing"},
        @{n=3;x=862;y=172;c="49D99A";k="document"},
        @{n=4;x=862;y=316;c="4F8BE8";k="shield"},
        @{n=5;x=542;y=316;c="49D99A";k="layers"},
        @{n=6;x=222;y=316;c="49D99A";k="database"},
        @{n=7;x=222;y=458;c="F0832F";k="cycle"},
        @{n=8;x=542;y=458;c="F0832F";k="hierarchy"},
        @{n=9;x=862;y=458;c="9B7BE0";k="target"}
    )

    foreach($d in $icons){
        $n=[int]$d.n;$x=[double]$d.x;$y=[double]$d.y;$c=[string]$d.c
        Plate $sl $x $y $c $n|Out-Null
        switch($d.k){
            "ship"{
                $h=$sl.Shapes.AddShape(7,$x+8,$y+12,31,7);NoFillLine $h $c 1.0;N $h $n "Hull"
                $r=$sl.Shapes.AddShape(1,$x+16,$y+7,15,6);NoFillLine $r $c 0.9;N $r $n "Deck"
                $l=L $sl ($x+23) ($y+7) ($x+23) ($y+2) $c 0.9 ("RefIcon{0}_Mast"-f$n)
                $l=L $sl ($x+20) ($y+4) ($x+26) ($y+4) $c 0.7 ("RefIcon{0}_MastBar"-f$n)
                foreach($xx in @(12,18,24,30,36)){$q=L $sl ($x+$xx) ($y+18) ($x+$xx+3) ($y+15) $c 0.55 ("RefIcon{0}_Grid{1}"-f$n,$xx)}
            }
            "drawing"{
                $s=$sl.Shapes.AddShape(1,$x+10,$y+3,29,16);NoFillLine $s $c 1.0;N $s $n "Sheet"
                $v=L $sl ($x+25) ($y+5) ($x+25) ($y+17) $c 0.55 ("RefIcon{0}_Divide"-f$n)
                foreach($yy in @(7,11,15)){$q=L $sl ($x+14) ($y+$yy) ($x+22) ($y+$yy) $c 0.65 ("RefIcon{0}_Dim{1}"-f$n,$yy)}
                $o=$sl.Shapes.AddShape(9,$x+29,$y+7,6,6);NoFillLine $o $c 0.65;N $o $n "Plan"
            }
            "document"{
                $s=$sl.Shapes.AddShape(1,$x+14,$y+2,22,18);NoFillLine $s $c 1.0;N $s $n "Page"
                $q=L $sl ($x+29) ($y+2) ($x+36) ($y+8) $c 0.8 ("RefIcon{0}_Fold1"-f$n)
                $q=L $sl ($x+29) ($y+2) ($x+29) ($y+8) $c 0.8 ("RefIcon{0}_Fold2"-f$n)
                $q=L $sl ($x+29) ($y+8) ($x+36) ($y+8) $c 0.8 ("RefIcon{0}_Fold3"-f$n)
                foreach($yy in @(8,12,16)){
                    $o=$sl.Shapes.AddShape(9,$x+17,$y+$yy,2.5,2.5);FillLine $o $c $c 0.2;N $o $n ("Bullet"+$yy)
                    $q=L $sl ($x+22) ($y+$yy+1) ($x+31) ($y+$yy+1) $c 0.55 ("RefIcon{0}_Text{1}"-f$n,$yy)
                }
            }
            "shield"{
                $s=$sl.Shapes.AddShape(9,$x+14,$y+2,22,18);NoFillLine $s $c 1.1;N $s $n "ApprovalSeal"
                $q=L $sl ($x+19) ($y+11) ($x+23) ($y+15) $c 1.3 ("RefIcon{0}_Check1"-f$n)
                $q=L $sl ($x+23) ($y+15) ($x+31) ($y+7) $c 1.3 ("RefIcon{0}_Check2"-f$n)
            }
            "layers"{
                foreach($off in @(0,5,10)){
                    $f=$sl.Shapes.AddShape(4,$x+13,$y+2+$off,23,8);NoFillLine $f $c 0.9;N $f $n ("Layer"+$off)
                }
            }
            "database"{
                $top=$sl.Shapes.AddShape(9,$x+14,$y+3,22,7);NoFillLine $top $c 1.0;N $top $n "Top"
                $bot=$sl.Shapes.AddShape(9,$x+14,$y+13,22,7);NoFillLine $bot $c 1.0;N $bot $n "Bottom"
                $q=L $sl ($x+14) ($y+6.5) ($x+14) ($y+16.5) $c 1.0 ("RefIcon{0}_Left"-f$n)
                $q=L $sl ($x+36) ($y+6.5) ($x+36) ($y+16.5) $c 1.0 ("RefIcon{0}_Right"-f$n)
                $q=L $sl ($x+15) ($y+11) ($x+35) ($y+11) $c 0.65 ("RefIcon{0}_Mid"-f$n)
            }
            "cycle"{
                $a=L $sl ($x+13) ($y+6) ($x+34) ($y+6) $c 1.1 ("RefIcon{0}_Cycle1"-f$n);$a.Line.EndArrowheadStyle=3
                $a=L $sl ($x+36) ($y+9) ($x+36) ($y+15) $c 1.1 ("RefIcon{0}_Cycle2"-f$n);$a.Line.EndArrowheadStyle=3
                $a=L $sl ($x+34) ($y+17) ($x+13) ($y+17) $c 1.1 ("RefIcon{0}_Cycle3"-f$n);$a.Line.EndArrowheadStyle=3
                $a=L $sl ($x+11) ($y+15) ($x+11) ($y+9) $c 1.1 ("RefIcon{0}_Cycle4"-f$n);$a.Line.EndArrowheadStyle=3
            }
            "hierarchy"{
                $root=$sl.Shapes.AddShape(1,$x+21,$y+2,8,6);NoFillLine $root $c 0.9;N $root $n "Root"
                $q=L $sl ($x+25) ($y+8) ($x+25) ($y+13) $c 0.8 ("RefIcon{0}_Stem"-f$n)
                $q=L $sl ($x+12) ($y+13) ($x+38) ($y+13) $c 0.8 ("RefIcon{0}_Branch"-f$n)
                foreach($xx in @(10,22,34)){
                    $q=L $sl ($x+$xx+4) ($y+13) ($x+$xx+4) ($y+16) $c 0.7 ("RefIcon{0}_Drop{1}"-f$n,$xx)
                    $r=$sl.Shapes.AddShape(1,$x+$xx,$y+16,8,4);NoFillLine $r $c 0.7;N $r $n ("Child"+$xx)
                }
            }
            "target"{
                foreach($sz in @(20,12,4)){
                    $o=$sl.Shapes.AddShape(9,$x+24-($sz/2),$y+11-($sz/2),$sz,$sz);NoFillLine $o $c 0.8;N $o $n ("Ring"+$sz)
                }
                $q=L $sl ($x+24) ($y+0) ($x+24) ($y+22) $c 0.55 ("RefIcon{0}_V"-f$n)
                $q=L $sl ($x+13) ($y+11) ($x+35) ($y+11) $c 0.55 ("RefIcon{0}_H"-f$n)
            }
        }
    }

    $pres.Save()
    $sl.Export($render,"PNG",1600,900)
}finally{
    if($pres){try{$pres.Close()}catch{}}
    if($ppt){try{if($ppt.Presentations.Count-eq0){$ppt.Quit()}}catch{}}
    [GC]::Collect();[GC]::WaitForPendingFinalizers()
}
Write-Output ("BACKUP="+$backup)
Write-Output ("RENDER="+$render)
