$ErrorActionPreference="Stop"
$dir="D:\이력서\AVEVA - Marine Principal Technical Support & Consultant – PLM SME, Busan\Proposal"
$deck=Join-Path $dir "Future_Industrial_PLM_Meeting_Deck_EN_Rebuilt_v2_VisualReviewed.pptx"
$img="C:\Users\namma\.claude\plm_slide_work\visualreview_current\ship_3d_icon_crop.png"
$stamp=Get-Date -Format "yyyyMMdd_HHmmss"
$backupDir=Join-Path $dir ("_backup_"+$stamp)
$backup=Join-Path $backupDir ("Future_Industrial_PLM_Meeting_Deck_EN_Rebuilt_v2_VisualReviewed_BEFORE_SHIP_IMAGE_"+$stamp+".pptx")
$render="C:\Users\namma\.claude\plm_slide_work\visualreview_current\slide18_ship_image.png"
New-Item -ItemType Directory -Force -Path $backupDir|Out-Null
Copy-Item -LiteralPath $deck -Destination $backup -Force

$ppt=$null;$pres=$null
try{
    $ppt=New-Object -ComObject PowerPoint.Application
    $pres=$ppt.Presentations.Open($deck,$false,$false,$false)
    $sl=$pres.Slides.Item(18)
    for($i=$sl.Shapes.Count;$i-ge1;$i--){
        $nm=$sl.Shapes.Item($i).Name
        if($nm-like"RefIcon1_*" -or $nm-eq"Step1ShipRaster"){$sl.Shapes.Item($i).Delete()}
    }
    # Insert the actual reference-image crop into the Step 1 card.
    $p=$sl.Shapes.AddPicture($img,0,-1,[single]222,[single]170,[single]48,[single]31)
    $p.Name="Step1ShipRaster"
    $pres.Save()
    $sl.Export($render,"PNG",1600,900)
}finally{
    if($pres){try{$pres.Close()}catch{}}
    if($ppt){try{if($ppt.Presentations.Count-eq0){$ppt.Quit()}}catch{}}
}
Write-Output ("BACKUP="+$backup)
Write-Output ("RENDER="+$render)
