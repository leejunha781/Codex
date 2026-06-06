$ErrorActionPreference = "Stop"

$deck = "D:\이력서\AVEVA - Marine Principal Technical Support & Consultant – PLM SME, Busan\Proposal\Future_Industrial_PLM_Meeting_Deck_EN_Rebuilt_v2_VisualReviewed.pptx"
$work = "C:\Users\namma\.claude\plm_slide_work\visualreview_current"
$backup = Join-Path $work "VisualReviewed_before_shape_text_merge.pptx"
$render = Join-Path $work "render_after_merge"
$log = Join-Path $work "shape_text_merge_log.txt"
Copy-Item -LiteralPath $deck -Destination $backup -Force
New-Item -ItemType Directory -Force -Path $render | Out-Null
Get-ChildItem -LiteralPath $render -Filter "*.png" -ErrorAction SilentlyContinue | Remove-Item -Force

function ShapeArea($s) { return [double]$s.Width * [double]$s.Height }
function ContainsMostly($outer,$inner) {
    $ix1 = [math]::Max([double]$outer.Left,[double]$inner.Left)
    $iy1 = [math]::Max([double]$outer.Top,[double]$inner.Top)
    $ix2 = [math]::Min([double]$outer.Left + [double]$outer.Width,[double]$inner.Left + [double]$inner.Width)
    $iy2 = [math]::Min([double]$outer.Top + [double]$outer.Height,[double]$inner.Top + [double]$inner.Height)
    if ($ix2 -le $ix1 -or $iy2 -le $iy1) { return 0.0 }
    return (($ix2-$ix1)*($iy2-$iy1)) / (ShapeArea $inner)
}
function ClampMargin($v) {
    if ($v -lt 0) { return [single]0 }
    return [single]$v
}
function CopyTextIntoShape($tb,$sh) {
    $srcTf = $tb.TextFrame
    $srcTr = $srcTf.TextRange
    $dstTf = $sh.TextFrame
    $dstTf.AutoSize = 0
    $dstTf.WordWrap = $srcTf.WordWrap
    try { $dstTf.VerticalAnchor = $srcTf.VerticalAnchor } catch {}
    try { $dstTf.Orientation = $srcTf.Orientation } catch {}

    $dstTf.MarginLeft   = ClampMargin (([double]$tb.Left - [double]$sh.Left) + [double]$srcTf.MarginLeft)
    $dstTf.MarginTop    = ClampMargin (([double]$tb.Top - [double]$sh.Top) + [double]$srcTf.MarginTop)
    $dstTf.MarginRight  = ClampMargin ((([double]$sh.Left + [double]$sh.Width) - ([double]$tb.Left + [double]$tb.Width)) + [double]$srcTf.MarginRight)
    $dstTf.MarginBottom = ClampMargin ((([double]$sh.Top + [double]$sh.Height) - ([double]$tb.Top + [double]$tb.Height)) + [double]$srcTf.MarginBottom)

    $dstTr = $dstTf.TextRange
    $dstTr.Text = $srcTr.Text
    try { $dstTr.Font.Name = $srcTr.Font.Name } catch {}
    try { $dstTr.Font.Size = [single]$srcTr.Font.Size } catch {}
    try { $dstTr.Font.Bold = [int]$srcTr.Font.Bold } catch {}
    try { $dstTr.Font.Italic = [int]$srcTr.Font.Italic } catch {}
    try { $dstTr.Font.Underline = [int]$srcTr.Font.Underline } catch {}
    try { $dstTr.Font.Color.RGB = [int]$srcTr.Font.Color.RGB } catch {}
    try { $dstTr.ParagraphFormat.Alignment = [int]$srcTr.ParagraphFormat.Alignment } catch {}
    try { $dstTr.ParagraphFormat.Bullet.Visible = [int]$srcTr.ParagraphFormat.Bullet.Visible } catch {}
}

$ppt = $null
$pres = $null
$sb = New-Object System.Text.StringBuilder
$merged = 0
try {
    $ppt = New-Object -ComObject PowerPoint.Application
    $pres = $ppt.Presentations.Open($deck,$false,$false,$false)
    foreach ($idx in @(14,15,17,18,19,21,22,24,25,26,28,30)) {
        $sl = $pres.Slides.Item($idx)
        $pairs = New-Object System.Collections.ArrayList
        $targetCounts = @{}
        foreach ($tb in $sl.Shapes) {
            if ($tb.Type -ne 17) { continue }
            try {
                if ($tb.HasTextFrame -ne -1 -or $tb.TextFrame.HasText -ne -1) { continue }
                $runs = $tb.TextFrame.TextRange.Runs().Count
                if ($runs -ne 1) { continue }
            } catch { continue }
            $best = $null
            $bestScore = 0
            foreach ($sh in $sl.Shapes) {
                if ($sh.Name -eq $tb.Name -or $sh.Type -eq 17 -or $sh.Type -eq 13 -or $sh.Type -eq 6) { continue }
                $hasText = $false
                try { $hasText = ($sh.HasTextFrame -eq -1 -and $sh.TextFrame.HasText -eq -1) } catch {}
                if ($hasText) { continue }
                $score = ContainsMostly $sh $tb
                if ($score -gt $bestScore -and (ShapeArea $sh) -lt ((ShapeArea $tb) * 3.0)) {
                    $best = $sh
                    $bestScore = $score
                }
            }
            if ($best -and $bestScore -ge 0.92) {
                [void]$pairs.Add(@($tb,$best,$bestScore))
                if (-not $targetCounts.ContainsKey($best.Name)) { $targetCounts[$best.Name] = 0 }
                $targetCounts[$best.Name]++
            }
        }

        # Merge only one-textbox-to-one-empty-shape pairs. Cards with separate title/body boxes remain untouched.
        foreach ($pair in @($pairs)) {
            $tb = $pair[0]
            $sh = $pair[1]
            if ($targetCounts[$sh.Name] -ne 1) { continue }
            $txt = ($tb.TextFrame.TextRange.Text -replace "[\r\n\v]"," ").Trim()
            CopyTextIntoShape $tb $sh
            $tb.Delete()
            $merged++
            [void]$sb.AppendLine(("[{0:D2}] {1} -> {2} :: {3}" -f $idx,$tb.Name,$sh.Name,$txt))
        }
    }
    $pres.Save()
    foreach ($idx in @(14,15,17,18,19,21,22,24,25,26,28,30,31,32,33)) {
        $pres.Slides.Item($idx).Export((Join-Path $render ("slide-{0:D2}.png" -f $idx)),"PNG",1280,720)
    }
}
finally {
    if ($pres) { try { $pres.Close() } catch {} }
    if ($ppt) { try { if ($ppt.Presentations.Count -eq 0) { $ppt.Quit() } } catch {} }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
[System.IO.File]::WriteAllText($log,$sb.ToString(),[System.Text.Encoding]::UTF8)
Write-Output ("MERGED=" + $merged)
Write-Output ("LOG=" + $log)
Write-Output ("RENDER=" + $render)
