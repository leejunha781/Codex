$ErrorActionPreference="Stop"
$deck="D:\이력서\AVEVA - Marine Principal Technical Support & Consultant – PLM SME, Busan\Proposal\Future_Industrial_PLM_Meeting_Deck_EN_Rebuilt_v2.pptx"
$reviewDir="C:\Users\namma\.claude\plm_slide_work\enhanced_full_deck_review_20260606"
$ppt=$null;$pres=$null
try{
    $ppt=New-Object -ComObject PowerPoint.Application
    $pres=$ppt.Presentations.Open($deck,0,0,0)
    $pres.Slides.Item(2).Shapes.Item("TextBox 24").TextFrame.TextRange.Text="PUBLIC SIGNALS"
    $pres.Slides.Item(2).Shapes.Item("TextBox 25").TextFrame.TextRange.Text="Official product and roadmap signals; confirmed versus directional"
    $pres.Slides.Item(6).Shapes.Item("TextBox 13").TextFrame.TextRange.Text="Public product + roadmap signals - official emphasis only"
    $pres.Slides.Item(6).Shapes.Item("TextBox 49").TextFrame.TextRange.Text="Treat public signals as directional, not guaranteed scope. Build adapters and an assurance graph that can absorb change without rewriting the platform."
    $pres.Save()
    foreach($sl in $pres.Slides){$sl.Export((Join-Path $reviewDir ("slide-{0:D2}.png" -f $sl.SlideIndex)),"PNG",1280,720)}
    Write-Output "SIGNAL_LABEL_FIX_DONE"
}finally{
    if($pres){try{$pres.Close()}catch{};try{[Runtime.InteropServices.Marshal]::ReleaseComObject($pres)|Out-Null}catch{}}
    if($ppt){try{if($ppt.Presentations.Count-eq 0){$ppt.Quit()}}catch{};try{[Runtime.InteropServices.Marshal]::ReleaseComObject($ppt)|Out-Null}catch{}}
    [GC]::Collect();[GC]::WaitForPendingFinalizers()
}
