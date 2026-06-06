$ErrorActionPreference = "Stop"
Add-Type -AssemblyName System.Drawing
$rdir = "C:\Users\namma\.claude\plm_slide_work\render_v2"
function Crop($file,$x,$y,$w,$h,$out){
  $src=New-Object System.Drawing.Bitmap($file)
  $r=New-Object System.Drawing.Rectangle($x,$y,$w,$h)
  $c=$src.Clone($r,$src.PixelFormat); $c.Save($out,[System.Drawing.Imaging.ImageFormat]::Png)
  $src.Dispose();$c.Dispose()
}
Crop "$rdir\img2_21.png" 0 0 1100 180 "$rdir\c21_header.png"
# slide 9 right panel: L558 T116 W366 H206 in 960 -> *1.6667
Crop "$rdir\img2_09.png" 920 185 620 360 "$rdir\c9_panel.png"
"cropped"
