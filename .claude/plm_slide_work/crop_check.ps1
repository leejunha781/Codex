$ErrorActionPreference = "Stop"
Add-Type -AssemblyName System.Drawing
$rdir = "C:\Users\namma\.claude\plm_slide_work\render_v2"
function Crop($file,$x,$y,$w,$h,$out){
  $src = New-Object System.Drawing.Bitmap($file)
  $rect = New-Object System.Drawing.Rectangle($x,$y,$w,$h)
  $crop = $src.Clone($rect, $src.PixelFormat)
  $crop.Save($out, [System.Drawing.Imaging.ImageFormat]::Png)
  $src.Dispose(); $crop.Dispose()
}
# slide 8 header (title+subtitle) : 1600x900 render
Crop "$rdir\img_edit08.png" 0 0 1000 200 "$rdir\c8_header.png"
# slide 16 VISUAL PROOF panel : L38 T236 W520 H293 in 960x540 -> *1.6667
Crop "$rdir\img_edit16.png" 55 385 880 500 "$rdir\c16_panel.png"
"cropped"
