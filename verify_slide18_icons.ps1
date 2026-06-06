$ErrorActionPreference="Stop"
$deck="D:\이력서\AVEVA - Marine Principal Technical Support & Consultant – PLM SME, Busan\Proposal\Future_Industrial_PLM_Meeting_Deck_EN_Rebuilt_v2_VisualReviewed.pptx"
$ppt=$null;$pres=$null
try{
    $ppt=New-Object -ComObject PowerPoint.Application
    $pres=$ppt.Presentations.Open($deck,$true,$false,$false)
    $sl=$pres.Slides.Item(18)
    $icons=0
    $refIcons=0
    foreach($sh in $sl.Shapes){
        if($sh.Name-like"HelperIcon*"){$icons++}
        if($sh.Name-like"RefIcon*"){$refIcons++}
    }
    Write-Output ("SLIDES="+$pres.Slides.Count)
    Write-Output ("SLIDE18_HELPER_SHAPES="+$icons)
    Write-Output ("SLIDE18_REFERENCE_ICON_SHAPES="+$refIcons)
    Write-Output ("SLIDE18_BACKGROUND_FILL_TYPE="+$sl.Background.Fill.Type)
}finally{
    if($pres){try{$pres.Close()}catch{}}
    if($ppt){try{if($ppt.Presentations.Count-eq0){$ppt.Quit()}}catch{}}
}
try{$fs=[IO.File]::Open($deck,'Open','ReadWrite','None');$fs.Close();Write-Output "LOCK=unlocked"}catch{Write-Output "LOCK=LOCKED"}
