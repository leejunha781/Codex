$ErrorActionPreference = 'Stop'

$path = 'C:\Users\namma\.claude\plm_slide_work\v6_future_direction_outline.txt'
$lines = Get-Content -LiteralPath $path
$markers = Select-String -LiteralPath $path -Pattern '^===== SLIDE '

foreach ($marker in $markers) {
    $slideNo = ($marker.Line -replace '[^0-9]', '')
    $next = $markers | Where-Object { $_.LineNumber -gt $marker.LineNumber } | Select-Object -First 1
    $endLine = if ($next) { $next.LineNumber - 2 } else { $lines.Count - 1 }
    $title = ''
    for ($i = $marker.LineNumber; $i -le $endLine; $i++) {
        $candidate = $lines[$i].Trim()
        if ($candidate.Length -eq 0) { continue }
        if ($candidate -match '^\d+$') { continue }
        if ($candidate -eq 'Future Industrial PLM') { continue }
        $title = $candidate
        break
    }
    "{0}. {1}" -f $slideNo, $title
}
