$ErrorActionPreference = "Stop"

$dir = "D:\이력서\AVEVA - Marine Principal Technical Support & Consultant – PLM SME, Busan\Proposal"
$src = Join-Path $dir "Future_Industrial_PLM_Meeting_Deck_EN_Rebuilt_v2.pptx"
$out = Join-Path $dir "Future_Industrial_PLM_Meeting_Deck_EN_Rebuilt_v2_VisualReviewed.pptx"
$stamp = Get-Date -Format "yyyyMMdd_HHmmss"
$backupDir = Join-Path $dir ("_backup_" + $stamp)
$backup = Join-Path $backupDir ("Future_Industrial_PLM_Meeting_Deck_EN_Rebuilt_v2_BACKUP_" + $stamp + ".pptx")
$work = "C:\Users\namma\.claude\plm_slide_work\visualreview_current"
$render = Join-Path $work "render_before"
$audit = Join-Path $work "audit.txt"

New-Item -ItemType Directory -Force -Path $backupDir | Out-Null
New-Item -ItemType Directory -Force -Path $render | Out-Null
Get-ChildItem -LiteralPath $render -Filter "*.png" -ErrorAction SilentlyContinue | Remove-Item -Force
Copy-Item -LiteralPath $src -Destination $backup -Force
Copy-Item -LiteralPath $src -Destination $out -Force

$ppt = $null
$pres = $null
$sb = New-Object System.Text.StringBuilder
try {
    $ppt = New-Object -ComObject PowerPoint.Application
    $pres = $ppt.Presentations.Open($out, $true, $false, $false)
    [void]$sb.AppendLine("DECK=" + $out)
    [void]$sb.AppendLine("SLIDES=" + $pres.Slides.Count)

    foreach ($sl in $pres.Slides) {
        $title = ""
        $tableCount = 0
        $pictureCount = 0
        $textBoxes = 0
        $autoshapesWithText = 0
        foreach ($sh in $sl.Shapes) {
            try { if ($sh.HasTable) { $tableCount++ } } catch {}
            if ($sh.Type -eq 13) { $pictureCount++ }
            try {
                if ($sh.HasTextFrame -eq -1 -and $sh.TextFrame.HasText -eq -1) {
                    $t = ($sh.TextFrame.TextRange.Text -replace "[\r\n\v]"," ").Trim()
                    if (-not $title -and $t) { $title = $t }
                    if ($sh.Type -eq 17) { $textBoxes++ } else { $autoshapesWithText++ }
                }
            } catch {}
        }
        [void]$sb.AppendLine(("[{0:D2}] tables={1} pictures={2} textboxes={3} shapesWithText={4} :: {5}" -f
            $sl.SlideIndex,$tableCount,$pictureCount,$textBoxes,$autoshapesWithText,$title))
        $sl.Export((Join-Path $render ("slide-{0:D2}.png" -f $sl.SlideIndex)), "PNG", 1280, 720)
    }
}
finally {
    if ($pres) { try { $pres.Close() } catch {} }
    if ($ppt) { try { if ($ppt.Presentations.Count -eq 0) { $ppt.Quit() } } catch {} }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

[System.IO.File]::WriteAllText($audit, $sb.ToString(), [System.Text.Encoding]::UTF8)
Write-Output ("BACKUP=" + $backup)
Write-Output ("WORKING_COPY=" + $out)
Write-Output ("AUDIT=" + $audit)
Write-Output ("RENDER=" + $render)
