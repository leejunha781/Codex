$ErrorActionPreference = "Stop"
$imgdir = "D:\이력서\AVEVA - Marine Principal Technical Support & Consultant – PLM SME, Busan\Proposal\plm_integrated_english_package\04_Presentation Image"
$dst = "C:\Users\namma\.claude\plm_slide_work"
Add-Type -AssemblyName System.Drawing
"==== 04_Presentation Image folder ===="
Get-ChildItem -Path $imgdir -File | Sort-Object Name | ForEach-Object {
  $b = New-Object System.Drawing.Bitmap($_.FullName)
  "{0,-55} {1,5}KB  {2}x{3}" -f $_.Name, [int]($_.Length/1KB), $b.Width, $b.Height
  $b.Dispose()
}
"==== copy 08 + 16 to work ===="
Copy-Item (Join-Path $imgdir "08_open_aveva_control_plane_generated.png") (Join-Path $dst "img08.png") -Force
Copy-Item (Join-Path $imgdir "16_continuous_naval_assurance_loop_generated.png") (Join-Path $dst "img16.png") -Force
"copied img08.png, img16.png"
