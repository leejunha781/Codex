$ErrorActionPreference = "Stop"

$deck = "D:\이력서\AVEVA - Marine Principal Technical Support & Consultant – PLM SME, Busan\Proposal\Future_Industrial_PLM_Meeting_Deck_EN_Rebuilt_v2_VisualReviewed.pptx"
$out = "C:\Users\namma\.claude\plm_slide_work\visualreview_current\shape_text_candidates.txt"

function Area($s) { return [double]$s.Width * [double]$s.Height }
function ContainsMostly($outer,$inner) {
    $ix1 = [math]::Max([double]$outer.Left,[double]$inner.Left)
    $iy1 = [math]::Max([double]$outer.Top,[double]$inner.Top)
    $ix2 = [math]::Min([double]$outer.Left + [double]$outer.Width,[double]$inner.Left + [double]$inner.Width)
    $iy2 = [math]::Min([double]$outer.Top + [double]$outer.Height,[double]$inner.Top + [double]$inner.Height)
    if ($ix2 -le $ix1 -or $iy2 -le $iy1) { return 0.0 }
    return (($ix2-$ix1)*($iy2-$iy1)) / (Area $inner)
}

$ppt = $null
$pres = $null
$sb = New-Object System.Text.StringBuilder
try {
    $ppt = New-Object -ComObject PowerPoint.Application
    $pres = $ppt.Presentations.Open($deck,$true,$false,$false)
    # Focus on the UI/process slides where separate text-over-shape construction is most likely.
    foreach ($idx in @(14,15,17,18,19,21,22,24,25,26,28,30)) {
        $sl = $pres.Slides.Item($idx)
        $candidates = 0
        foreach ($tb in $sl.Shapes) {
            if ($tb.Type -ne 17) { continue }
            $txt = ""
            try {
                if ($tb.HasTextFrame -ne -1 -or $tb.TextFrame.HasText -ne -1) { continue }
                $txt = ($tb.TextFrame.TextRange.Text -replace "[\r\n\v]"," ").Trim()
            } catch { continue }
            if (-not $txt) { continue }
            $best = $null
            $bestScore = 0
            foreach ($sh in $sl.Shapes) {
                if ($sh.Name -eq $tb.Name -or $sh.Type -eq 17 -or $sh.Type -eq 13 -or $sh.Type -eq 6) { continue }
                $hasText = $false
                try { $hasText = ($sh.HasTextFrame -eq -1 -and $sh.TextFrame.HasText -eq -1) } catch {}
                if ($hasText) { continue }
                $score = ContainsMostly $sh $tb
                if ($score -gt $bestScore -and (Area $sh) -lt ((Area $tb) * 3.0)) {
                    $best = $sh
                    $bestScore = $score
                }
            }
            if ($best -and $bestScore -ge 0.92) {
                $candidates++
                $short = $txt
                if ($short.Length -gt 80) { $short = $short.Substring(0,80) + "..." }
                [void]$sb.AppendLine(("[{0:D2}] textbox='{1}' -> shape='{2}' score={3:N2} :: {4}" -f
                    $sl.SlideIndex,$tb.Name,$best.Name,$bestScore,$short))
            }
        }
        if ($candidates -gt 0) { [void]$sb.AppendLine(("  slide candidates=" + $candidates)) }
    }
}
finally {
    if ($pres) { try { $pres.Close() } catch {} }
    if ($ppt) { try { if ($ppt.Presentations.Count -eq 0) { $ppt.Quit() } } catch {} }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
[System.IO.File]::WriteAllText($out,$sb.ToString(),[System.Text.Encoding]::UTF8)
Write-Output ("OUT=" + $out)
