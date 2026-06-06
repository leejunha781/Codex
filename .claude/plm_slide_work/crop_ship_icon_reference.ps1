$ErrorActionPreference="Stop"
Add-Type -AssemblyName System.Drawing

$dir="D:\이력서\AVEVA - Marine Principal Technical Support & Consultant – PLM SME, Busan\Proposal"
$src=Join-Path $dir "plm_integrated_english_package\04_Presentation Image\a_wide_high_resolution_futuristic_infographic_d.png"
$out="C:\Users\namma\.claude\plm_slide_work\visualreview_current\ship_3d_icon_crop.png"

$img=[System.Drawing.Bitmap]::FromFile($src)
try{
    # Crop tightly around the ship symbol inside the first process tile, excluding label text.
    $rect=New-Object System.Drawing.Rectangle(43,392,113,73)
    $crop=$img.Clone($rect,$img.PixelFormat)
    try{$crop.Save($out,[System.Drawing.Imaging.ImageFormat]::Png)}finally{$crop.Dispose()}
}finally{$img.Dispose()}
Write-Output ("OUT="+$out)
