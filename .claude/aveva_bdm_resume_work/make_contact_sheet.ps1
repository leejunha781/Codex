$ErrorActionPreference = "Stop"

$renderDir = "C:\Users\namma\.claude\aveva_bdm_resume_work\bdm_render"
$outPath = "C:\Users\namma\.claude\aveva_bdm_resume_work\bdm_contact_sheet.png"
$files = Get-ChildItem -LiteralPath $renderDir -Filter "page-*.png" | Sort-Object Name

Add-Type -AssemblyName System.Drawing

$thumbW = 320
$thumbH = 414
$cols = 3
$rows = [int][Math]::Ceiling($files.Count / $cols)
$labelH = 28
$pad = 18
$sheetW = ($cols * $thumbW) + (($cols + 1) * $pad)
$sheetH = ($rows * ($thumbH + $labelH)) + (($rows + 1) * $pad)

$bitmap = New-Object System.Drawing.Bitmap $sheetW, $sheetH
$graphics = [System.Drawing.Graphics]::FromImage($bitmap)
$graphics.Clear([System.Drawing.Color]::White)
$graphics.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
$font = New-Object System.Drawing.Font "Arial", 13, ([System.Drawing.FontStyle]::Bold)
$brush = [System.Drawing.Brushes]::Black
$pen = New-Object System.Drawing.Pen ([System.Drawing.Color]::LightGray), 1

for ($i = 0; $i -lt $files.Count; $i++) {
    $col = $i % $cols
    $row = [int][Math]::Floor($i / $cols)
    $x = $pad + ($col * ($thumbW + $pad))
    $y = $pad + ($row * ($thumbH + $labelH + $pad))
    $img = [System.Drawing.Image]::FromFile($files[$i].FullName)
    try {
        $graphics.DrawImage($img, $x, ($y + $labelH), $thumbW, $thumbH)
        $graphics.DrawRectangle($pen, $x, ($y + $labelH), $thumbW, $thumbH)
        $label = "Page " + ($i + 1)
        $graphics.DrawString($label, $font, $brush, $x, $y)
    } finally {
        $img.Dispose()
    }
}

$bitmap.Save($outPath, [System.Drawing.Imaging.ImageFormat]::Png)
$graphics.Dispose()
$bitmap.Dispose()
$font.Dispose()
$pen.Dispose()

$outPath
