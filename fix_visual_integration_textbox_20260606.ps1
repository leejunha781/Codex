$ErrorActionPreference="Stop"
$deck="D:\이력서\AVEVA - Marine Principal Technical Support & Consultant – PLM SME, Busan\Proposal\Future_Industrial_PLM_Meeting_Deck_EN_Rebuilt_v2.pptx"
$reviewDir="C:\Users\namma\.claude\plm_slide_work\visual_integration_review_20260606"
$ppt=$null;$pres=$null
try{
    $ppt=New-Object -ComObject PowerPoint.Application
    $pres=$ppt.Presentations.Open($deck,0,0,0)
    $sh=$pres.Slides.Item(30).Shapes.Item("TextBox 14")
    $sh.Height=[single]50
    $pres.Save()
    foreach($sl in $pres.Slides){$sl.Export((Join-Path $reviewDir ("slide-{0:D2}.png" -f $sl.SlideIndex)),"PNG",1280,720)}
    Write-Output "TEXTBOX_FIX_DONE"
}finally{
    if($pres){try{$pres.Close()}catch{};try{[Runtime.InteropServices.Marshal]::ReleaseComObject($pres)|Out-Null}catch{}}
    if($ppt){try{if($ppt.Presentations.Count-eq0){$ppt.Quit()}}catch{};try{[Runtime.InteropServices.Marshal]::ReleaseComObject($ppt)|Out-Null}catch{}}
    [GC]::Collect();[GC]::WaitForPendingFinalizers()
}
