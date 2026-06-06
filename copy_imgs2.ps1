$ErrorActionPreference = "Stop"
$prop = "D:\이력서\AVEVA - Marine Principal Technical Support & Consultant – PLM SME, Busan\Proposal"
$dst = "C:\Users\namma\.claude\plm_slide_work"
"prop exists: " + (Test-Path $prop)
$map = @{
  "08_open_aveva_control_plane_generated.png" = "img08.png"
  "16_continuous_naval_assurance_loop_generated.png" = "img16.png"
  "06_a_highly_detailed_clean_futuristic_infographic.png" = "img06.png"
  "07_additional_architecture_visual.png" = "img07.png"
  "05_a_high_detail_isometric_3d_infographic_scene_a_la.png" = "img05.png"
}
Add-Type -AssemblyName System.Drawing
foreach($k in $map.Keys){
  $src = Join-Path $prop $k
  if(Test-Path $src){
    $out = Join-Path $dst $map[$k]
    Copy-Item $src $out -Force
    $b = New-Object System.Drawing.Bitmap($out)
    "{0,-10} <- {1}  ({2} x {3}px, ratio {4:N3})" -f $map[$k],$k,$b.Width,$b.Height,($b.Width/$b.Height)
    $b.Dispose()
  } else { "MISSING: $k" }
}
