$ErrorActionPreference = "Stop"
$v2 = "D:\이력서\AVEVA - Marine Principal Technical Support & Consultant – PLM SME, Busan\Proposal\Future_Industrial_PLM_Meeting_Deck_EN_Rebuilt_v2.pptx"
$name = "Future_Industrial_PLM_Meeting_Deck_EN_Rebuilt_v2.pptx"
$outDir = "C:\Users\namma\.claude\plm_slide_work\review_png8"
function Rgb([string]$hex){ $hex=$hex.TrimStart('#'); [int]([Convert]::ToInt32($hex.Substring(0,2),16)+[Convert]::ToInt32($hex.Substring(2,2),16)*256+[Convert]::ToInt32($hex.Substring(4,2),16)*65536) }
function Retry($sb){ for($a=0;$a -lt 6;$a++){ try{ return (& $sb) }catch{ Start-Sleep -Milliseconds 80 } }; return (& $sb) }
$ppt = New-Object -ComObject PowerPoint.Application
foreach($p in @($ppt.Presentations)){ if($p.Name -eq $name){ try{ $p.Saved=$true; $p.Close() }catch{} } }
$pres=$ppt.Presentations.Open($v2,0,0,0)
function FindLike($pres,$prefix){ foreach($sl in $pres.Slides){ foreach($sh in $sl.Shapes){ try{ if($sh.HasTextFrame -eq -1 -and $sh.TextFrame.HasText -eq -1){ if($sh.TextFrame.TextRange.Text.Trim().StartsWith($prefix)){ return $sl } } }catch{} } }; return $null }

$s = FindLike $pres "Recommended decision"
Write-Output ("s23 idx=" + $s.SlideIndex)
$script:slide = $s

# add a dark scrim over the left text region (sits above the photo)
$scrim = $script:slide.Shapes.AddShape(5,[single]18,[single]18,[single]600,[single]504)
try{$scrim.Adjustments.Item(1)=[single]0.03}catch{}
$scrim.Fill.Solid(); Retry{$scrim.Fill.ForeColor.RGB=[int](Rgb "05090F")}; $scrim.Fill.Transparency=[single]0.28
$scrim.Line.Visible=0; try{$scrim.Shadow.Visible=0}catch{}

# bring all text shapes to front so they sit above the scrim
foreach($sh in @($script:slide.Shapes)){
  try{
    if($sh.HasTextFrame -eq -1 -and $sh.TextFrame.HasText -eq -1){ $sh.ZOrder(0) }  # msoBringToFront
  }catch{}
}

$pres.Save()
$s.Export("$outDir\s23_scrim.png","PNG",1280,720)
$pres.Close()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($pres) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($ppt)  | Out-Null
[GC]::Collect(); [GC]::WaitForPendingFinalizers()
Write-Output "BUILD8_S23_DONE"
