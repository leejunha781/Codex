$ErrorActionPreference = "Stop"
Add-Type -AssemblyName System.Drawing

$dir = if ($env:CONTACT_DIR) { $env:CONTACT_DIR } else { "C:\Users\namma\.claude\plm_slide_work\visualreview_current\render_before" }
$out = if ($env:CONTACT_OUT) { $env:CONTACT_OUT } else { "C:\Users\namma\.claude\plm_slide_work\visualreview_current\contact_sheet_before.png" }
$files = Get-ChildItem -LiteralPath $dir -Filter "slide-*.png" | Sort-Object Name
$cols = [int]4
$thumbW = [int]320
$thumbH = [int]180
$labelH = [int]26
$rows = [int][math]::Ceiling(([double]$files.Count / [double]$cols))
$sheetW = [int]($cols * $thumbW)
$sheetH = [int]($rows * ($thumbH + $labelH))
$sheet = New-Object System.Drawing.Bitmap($sheetW,$sheetH)
$g = [System.Drawing.Graphics]::FromImage($sheet)
$g.Clear([System.Drawing.Color]::FromArgb(16,27,44))
$font = New-Object System.Drawing.Font("Arial",10,[System.Drawing.FontStyle]::Bold)
$brush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::White)
for ($i=0; $i -lt $files.Count; $i++) {
    $x = ($i % $cols) * $thumbW
    $y = [math]::Floor($i / $cols) * ($thumbH+$labelH)
    $img = [System.Drawing.Image]::FromFile($files[$i].FullName)
    $g.DrawImage($img,$x,$y,$thumbW,$thumbH)
    $g.DrawString($files[$i].BaseName,$font,$brush,$x+8,$y+$thumbH+4)
    $img.Dispose()
}
$sheet.Save($out,[System.Drawing.Imaging.ImageFormat]::Png)
$brush.Dispose(); $font.Dispose(); $g.Dispose(); $sheet.Dispose()
Write-Output ("OUT=" + $out)
