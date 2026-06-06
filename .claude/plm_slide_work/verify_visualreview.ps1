$ErrorActionPreference = "Stop"

$deck = "D:\이력서\AVEVA - Marine Principal Technical Support & Consultant – PLM SME, Busan\Proposal\Future_Industrial_PLM_Meeting_Deck_EN_Rebuilt_v2_VisualReviewed.pptx"
$work = "C:\Users\namma\.claude\plm_slide_work\visualreview_current"
$render = Join-Path $work "render_final"
$report = Join-Path $work "verification.txt"
New-Item -ItemType Directory -Force -Path $render | Out-Null
Get-ChildItem -LiteralPath $render -Filter "*.png" -ErrorAction SilentlyContinue | Remove-Item -Force

$ppt = $null
$pres = $null
$sb = New-Object System.Text.StringBuilder
$tableCount = 0
$pictureCount = 0
$textBoxCount = 0
$shapeTextCount = 0
$overflow = New-Object System.Collections.ArrayList
try {
    $ppt = New-Object -ComObject PowerPoint.Application
    $pres = $ppt.Presentations.Open($deck,$true,$false,$false)
    [void]$sb.AppendLine("SLIDES=" + $pres.Slides.Count)
    foreach ($sl in $pres.Slides) {
        foreach ($sh in $sl.Shapes) {
            try { if ($sh.HasTable) { $tableCount++ } } catch {}
            if ($sh.Type -eq 13) { $pictureCount++ }
            try {
                if ($sh.HasTextFrame -eq -1 -and $sh.TextFrame.HasText -eq -1) {
                    if ($sh.Type -eq 17) { $textBoxCount++ } else { $shapeTextCount++ }
                    $avail = [double]$sh.Height - [double]$sh.TextFrame.MarginTop - [double]$sh.TextFrame.MarginBottom
                    $bound = [double]$sh.TextFrame.TextRange.BoundHeight
                    if ($bound -gt ($avail + 2.5)) {
                        $txt = ($sh.TextFrame.TextRange.Text -replace "[\r\n\v]"," ").Trim()
                        if ($txt.Length -gt 60) { $txt = $txt.Substring(0,60) + "..." }
                        [void]$overflow.Add(("[{0:D2}] {1} bound={2:N1} avail={3:N1} :: {4}" -f $sl.SlideIndex,$sh.Name,$bound,$avail,$txt))
                    }
                }
            } catch {}
        }
        $sl.Export((Join-Path $render ("slide-{0:D2}.png" -f $sl.SlideIndex)),"PNG",1280,720)
    }
}
finally {
    if ($pres) { try { $pres.Close() } catch {} }
    if ($ppt) { try { if ($ppt.Presentations.Count -eq 0) { $ppt.Quit() } } catch {} }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

[void]$sb.AppendLine("TABLES=" + $tableCount)
[void]$sb.AppendLine("PICTURES=" + $pictureCount + " (excludes slide-background images)")
[void]$sb.AppendLine("TEXTBOXES=" + $textBoxCount)
[void]$sb.AppendLine("AUTOSHAPES_WITH_TEXT=" + $shapeTextCount)
[void]$sb.AppendLine("POTENTIAL_OVERFLOW=" + $overflow.Count)
foreach ($x in $overflow) { [void]$sb.AppendLine($x) }
[System.IO.File]::WriteAllText($report,$sb.ToString(),[System.Text.Encoding]::UTF8)

Add-Type -AssemblyName System.IO.Compression.FileSystem
$z = [System.IO.Compression.ZipFile]::OpenRead($deck)
try {
    $zipSlides = ($z.Entries | Where-Object { $_.FullName -match '^ppt/slides/slide[0-9]+\.xml$' }).Count
} finally { $z.Dispose() }

$lock = "unlocked"
try { $fs=[System.IO.File]::Open($deck,'Open','ReadWrite','None'); $fs.Close() } catch { $lock="LOCKED" }
Write-Output ("PPTX_SLIDES=" + $zipSlides)
Write-Output ("LOCK=" + $lock)
Write-Output ("REPORT=" + $report)
Write-Output ("RENDER=" + $render)
