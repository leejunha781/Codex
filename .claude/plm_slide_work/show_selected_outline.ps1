$ErrorActionPreference = 'Stop'

$path = 'C:\Users\namma\.claude\plm_slide_work\v6_future_direction_outline.txt'
$slides = @(1, 2, 13, 14, 15, 16, 17, 28, 29, 30, 31, 32, 33, 34, 35, 36, 37)
$lines = Get-Content -LiteralPath $path
$markers = Select-String -LiteralPath $path -Pattern '^===== SLIDE '

foreach ($n in $slides) {
    $marker = $markers | Where-Object { $_.Line -eq "===== SLIDE $n =====" } | Select-Object -First 1
    if ($null -eq $marker) { continue }

    $next = $markers | Where-Object { $_.LineNumber -gt $marker.LineNumber } | Select-Object -First 1
    $endLine = if ($next) { $next.LineNumber - 2 } else { $lines.Count - 1 }
    $endLine = [Math]::Min($endLine, $marker.LineNumber + 35)

    $lines[($marker.LineNumber - 1)..$endLine]
    ''
}
