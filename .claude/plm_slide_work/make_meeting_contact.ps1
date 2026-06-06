$ErrorActionPreference="Stop"
Add-Type -AssemblyName System.Drawing
$dir="C:\Users\namma\.claude\plm_slide_work\meeting_current\render_before"
$out="C:\Users\namma\.claude\plm_slide_work\meeting_current\contact_sheet.png"
$fs=Get-ChildItem -LiteralPath $dir -Filter "slide-*.png"|Sort-Object Name
$cols=4;$tw=320;$th=180;$lh=24;$rows=[int][math]::Ceiling($fs.Count/$cols)
$bmp=New-Object Drawing.Bitmap([int]($cols*$tw),[int]($rows*($th+$lh)))
$g=[Drawing.Graphics]::FromImage($bmp);$g.Clear([Drawing.Color]::FromArgb(12,25,42))
$font=New-Object Drawing.Font("Arial",10,[Drawing.FontStyle]::Bold);$brush=[Drawing.Brushes]::White
for($i=0;$i-lt$fs.Count;$i++){
 $x=($i%$cols)*$tw;$y=[math]::Floor($i/$cols)*($th+$lh)
 $im=[Drawing.Image]::FromFile($fs[$i].FullName);$g.DrawImage($im,$x,$y,$tw,$th);$im.Dispose()
 $g.DrawString($fs[$i].BaseName,$font,$brush,$x+6,$y+$th+3)
}
$bmp.Save($out,[Drawing.Imaging.ImageFormat]::Png);$font.Dispose();$g.Dispose();$bmp.Dispose()
Write-Output("OUT="+$out)
