$dir = Get-ChildItem "D:\이력서" -Directory | Where-Object {$_.Name -like "AVEVA*PLM SME*"} | Select-Object -First 1
$file = Join-Path $dir.FullName "Proposal\Future_Industrial_PLM_Meeting_Deck_EN_Hybrid_Safety_Final_v4.pptx"
Write-Output "EXISTS: $(Test-Path $file)  SIZE: $((Get-Item $file).Length)  MODIFIED: $((Get-Item $file).LastWriteTime)"

$ppt = New-Object -ComObject PowerPoint.Application
$ppt.Visible = $false
try {
    $pres = $ppt.Presentations.Open($file, $true, $false, $false)
    $n = $pres.Slides.Count
    Write-Output "SLIDES: $n"
    $start = [Math]::Max(1, $n - 4)
    for ($i = $start; $i -le $n; $i++) {
        $s = $pres.Slides.Item($i)
        $txt = ""
        foreach ($sh in $s.Shapes) {
            try {
                if ($sh.HasTextFrame -eq -1) {
                    $raw = $sh.TextFrame.TextRange.Text
                    $short = $raw.Substring(0, [Math]::Min(80, $raw.Length))
                    $txt += " | " + $short
                }
            } catch {}
        }
        $disp = $txt.Substring(0, [Math]::Min(200, $txt.Length))
        Write-Output "s${i}: ${disp}"
    }
    $pres.Close()
} finally {
    if ($ppt.Presentations.Count -eq 0) { $ppt.Quit() }
}
