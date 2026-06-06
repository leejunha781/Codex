$file = "D:\이력서\AVEVA - Marine Principal Technical Support & Consultant – PLM SME, Busan\Proposal\Future_Industrial_PLM_Meeting_Deck_EN_Hybrid_Safety_Final_v4.pptx"
$outDir = "C:\Users\namma\.claude\plm_slide_work\v4_visual_slots"
New-Item -ItemType Directory -Force -Path $outDir | Out-Null

$pp = $null
$pres = $null
$openedHere = $false

function TextOfShape($sh) {
    $parts = New-Object System.Collections.Generic.List[string]
    try {
        if ($sh.TextFrame.HasText -eq -1) {
            $txt = ($sh.TextFrame.TextRange.Text -replace "[\r\n]+", " | ").Trim()
            if ($txt.Length -gt 0) { $parts.Add($txt) }
        }
    } catch {}
    try {
        if ($sh.HasTable -eq -1) {
            for ($r = 1; $r -le $sh.Table.Rows.Count; $r++) {
                $cells = @()
                for ($c = 1; $c -le $sh.Table.Columns.Count; $c++) {
                    try { $cells += (($sh.Table.Cell($r,$c).Shape.TextFrame.TextRange.Text -replace "[\r\n]+", " ").Trim()) } catch { $cells += "" }
                }
                $parts.Add("TBL: " + ($cells -join " | "))
            }
        }
    } catch {}
    try {
        if ($sh.Type -eq 6) {
            for ($i = 1; $i -le $sh.GroupItems.Count; $i++) {
                $sub = TextOfShape $sh.GroupItems.Item($i)
                if ($sub.Length -gt 0) { $parts.Add($sub) }
            }
        }
    } catch {}
    return ($parts -join " / ")
}

try {
    try {
        $pp = [Runtime.InteropServices.Marshal]::GetActiveObject("PowerPoint.Application")
    } catch {
        $pp = New-Object -ComObject PowerPoint.Application
        $openedHere = $true
    }
    $pp.Visible = $true
    foreach ($p in $pp.Presentations) {
        if ($p.FullName -eq $file) { $pres = $p; break }
    }
    if (-not $pres) {
        $pres = $pp.Presentations.Open($file, $false, $false, $false)
        $openedHere = $true
    }

    $slides = @(13,15,16,17,19,22,30,31)
    $report = New-Object System.Text.StringBuilder
    [void]$report.AppendLine("FILE: $file")
    [void]$report.AppendLine("SLIDES: $($pres.Slides.Count)")
    foreach ($n in $slides) {
        $sl = $pres.Slides.Item($n)
        [void]$report.AppendLine("")
        [void]$report.AppendLine("=== SLIDE $n ===")
        $picCount = 0
        foreach ($sh in $sl.Shapes) {
            $isPic = $false
            try {
                if ($sh.Type -eq 13 -or $sh.Type -eq 11) { $isPic = $true }
            } catch {}
            if ($isPic) { $picCount++ }
            $txt = TextOfShape $sh
            if ($isPic -or $txt.Length -gt 0) {
                $line = "{0,-28} L={1,6:N1} T={2,6:N1} W={3,6:N1} H={4,6:N1} {5}" -f $sh.Name, $sh.Left, $sh.Top, $sh.Width, $sh.Height, $txt
                [void]$report.AppendLine($line)
            }
        }
        [void]$report.AppendLine("PICTURES=$picCount")
        $png = Join-Path $outDir ("slide{0:D2}.png" -f $n)
        $sl.Export($png, "PNG", 1600, 900)
        [void]$report.AppendLine("PNG=$png")
    }
    $reportPath = Join-Path $outDir "audit_report.txt"
    [System.IO.File]::WriteAllText($reportPath, $report.ToString(), [System.Text.Encoding]::UTF8)
    Write-Output "REPORT=$reportPath"
} finally {
    if ($pres -and $openedHere) { $pres.Close() }
    if ($pp -and $openedHere) {
        try { if ($pp.Presentations.Count -eq 0) { $pp.Quit() } } catch {}
    }
}
