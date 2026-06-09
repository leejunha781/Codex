$ErrorActionPreference = "Stop"

$dir   = "D:\이력서\AVEVA - Marine Principal Technical Support & Consultant – PLM SME, Busan\Proposal\"
$v17   = $dir + "Future_Industrial_PLM_Meeting_Deck_EN_V17.pptx"
$v18   = $dir + "Future_Industrial_PLM_Meeting_Deck_EN_V18.pptx"
$alias = $dir + "Future_Industrial_PLM_Meeting_Deck_EN.pptx"
$final = $dir + "Future_Industrial_PLM_Meeting_Deck_EN_Final.pptx"
$fpdf  = $dir + "Future_Industrial_PLM_Meeting_Deck_EN_Final.pdf"
$bk    = $dir + "_backup_20260609_v18_slide2_mapping\"

New-Item -ItemType Directory -Force -Path $bk | Out-Null
foreach ($f in @($v17,$alias,$final)) { if (Test-Path $f) { Copy-Item $f -Destination $bk -Force } }
Copy-Item $v17 -Destination $v18 -Force

$nd=[char]0x2013; $md=[char]0x00B7; $ar=[char]0x2192
$navy=3678739; $panel=4205084; $cyan=15646772; $amber=3846888; $green=9224777
$white=16777215; $ltgray=15062727; $muted=12626067; $violet=16419751
$blue=16155195; $magenta=11161814; $teal=12571693; $orange=1471481; $rose=11956980
$titleFont="Aptos Display"; $bodyFont="Aptos"

Get-Process POWERPNT -ErrorAction SilentlyContinue | Stop-Process -Force
Start-Sleep -Milliseconds 700
foreach($n in @("StartupItems","DocumentRecovery")){ $k="HKCU:\Software\Microsoft\Office\16.0\PowerPoint\Resiliency\$n"; if(Test-Path $k){ Remove-Item $k -Recurse -Force } }
$null = Start-Process "POWERPNT.EXE" -PassThru
Start-Sleep -Seconds 4
$pp=$null
for($a=0;$a -lt 6 -and $pp -eq $null;$a++){ try{$pp=New-Object -ComObject PowerPoint.Application}catch{Start-Sleep -Milliseconds 1200} }
if($pp -eq $null){ throw "PowerPoint COM did not start" }
try { $pp.Visible=$true } catch {}

$pres = $pp.Presentations.Open($v18,$false,$false,$false)
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
function Prep-Slide($sl){
    foreach($sh in @($sl.Shapes)){ $sh.Delete() }
    $sl.FollowMasterBackground=$false
    $sl.Background.Fill.Solid(); $sl.Background.Fill.ForeColor.RGB=$navy
    return $sl
}

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

# precompute Lead -> audience labels (consultant role explicit + standardized tokens + mapping tweaks)
$H5="Lead " + $ar + " audience"
$L0="Facilitator " + $ar + " All"
$L1="Facilitator " + $md + " Consultant (PLM SME) " + $ar + " All"
$L2="Consultant (PLM SME) " + $ar + " Dev + Planning"
$L3="Consultant (PLM SME) " + $ar + " Planning (Dev on AI)"
$L4="Consultant (PLM SME) + Architect " + $ar + " Dev (Planning on scope)"
$L5="Consultant (PLM SME) " + $ar + " Dev + Planning"
$L6="Consultant (PLM SME) " + $ar + " Planning"
$L7="Facilitator " + $ar + " All"
$L8="Facilitator " + $ar + " All"

$rows=@(
 @("#","Min","Agenda block","Deck",$H5,"Outcome to capture / decide"),
 @("0","5","Open & decision objective","p.1",$L0,"Confirm the exact approval requested today"),
 @("1","5",("How we'll decide " + $nd + " procedure"),("p.4" + $nd + "6"),$L1,"Agree decision flow; set dev vs planning lenses"),
 @("2","10","Current state & why now",("p.7" + $nd + "10"),$L2,"Accept that the gap and urgency are real"),
 @("3","15","Winning ideas / strategy",("p.11" + $nd + "19"),$L3,"Agree win-above position + the AI stance"),
 @("4","20","Proposal & architecture",("p.20" + $nd + "30"),$L4,"Confirm hybrid architecture is feasible & bounded"),
 @("5","10",("Delivery " + $nd + " WBS & KPI gates"),("p.31" + $nd + "33"),$L5,"Agree the 26-wk staged plan + KPI acceptance"),
 @("6","5","Future-state payoff","p.34",$L6,"Judge the competitive upside credible"),
 @("7","10",("Review " + $nd + " ownership, risks, evidence"),("p.35" + $nd + "37"),$L7,"Assign owners; accept risks & assumptions"),
 @("8","10","Decision & close",("p.38" + $nd + "39"),$L8,"Record decision, scope, owners, first scenario")
)
$acc=@($cyan,$green,$blue,$magenta,$amber,$teal,$orange,$violet,$rose)

$tLeft=40; $tTop=150; $tWidth=($W-80); $nRows=$rows.Count; $nCols=6
$tblShape=$s.Shapes.AddTable($nRows,$nCols,[single]$tLeft,[single]$tTop,[single]$tWidth,[single]360)
$tbl=$tblShape.Table
$frac=@(0.038,0.050,0.220,0.072,0.255,0.365)
for($c=1;$c -le $nCols;$c++){ $tbl.Columns.Item($c).Width=[single]($tWidth*$frac[$c-1]) }
$aligns=@(2,2,1,2,1,1)
for($r=1;$r -le $nRows;$r++){
    $tbl.Rows.Item($r).Height=[single]36
    for($c=1;$c -le $nCols;$c++){
        $cell=$tbl.Cell($r,$c)
        $cell.Shape.TextFrame.VerticalAnchor=3
        $cell.Shape.TextFrame.MarginLeft=[single]6; $cell.Shape.TextFrame.MarginRight=[single]6
        $cell.Shape.TextFrame.MarginTop=[single]1; $cell.Shape.TextFrame.MarginBottom=[single]1
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

$legend="Consultant (PLM SME) = Principal Technical Support & Consultant (presenter)   " + $md + "   Facilitator = chair / decision owner   " + $md + "   Architect = Solution Architect   " + $md + "   Glossary p.40" + $nd + "42 = on-demand reference"
Add-Box $s 40 514 ($W-80) 16 $legend 8.5 0 $muted $bodyFont 1 | Out-Null

# bump cover revision -> V18
try { $pres.Slides.Item(1).Shapes.Item("RevisionStamp").TextFrame.TextRange.Text = "Rev. V18 " + $md + " 2026-06-09" } catch {}

$pres.Save()
$pres.SaveAs($alias)
$pres.SaveAs($final)
$pres.SaveAs($fpdf, 32)
$cnt=$pres.Slides.Count
$pres.Close()
$pp.Quit()
Write-Output ("V18 saved; slides=" + $cnt)
