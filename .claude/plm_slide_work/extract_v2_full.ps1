$ErrorActionPreference = "Stop"
$path = "D:\이력서\AVEVA - Marine Principal Technical Support & Consultant – PLM SME, Busan\Proposal\Future_Industrial_PLM_Meeting_Deck_EN_Rebuilt_v2.pptx"

$ppt = New-Object -ComObject PowerPoint.Application
try {
    $pres = $ppt.Presentations.Open($path, $true, $false, $false)  # ReadOnly, Untitled=false, WithWindow=false
    $sb = New-Object System.Text.StringBuilder

    function Dump-Shape($shp, $indent) {
        $pad = "  " * $indent
        $name = $shp.Name
        $stype = $shp.Type
        $hasText = $false
        try { $hasText = ($shp.HasTextFrame -and $shp.TextFrame.HasText) } catch {}
        $line = "$pad[shape] name='$name' type=$stype"
        # picture?
        if ($stype -eq 13) { $line += " <PICTURE>" }      # msoPicture
        if ($stype -eq 6)  { $line += " <GROUP>" }         # msoGroup
        try { $line += " pos=($([math]::Round($shp.Left)),$([math]::Round($shp.Top))) size=($([math]::Round($shp.Width))x$([math]::Round($shp.Height)))" } catch {}
        [void]$sb.AppendLine($line)
        if ($hasText) {
            $t = $shp.TextFrame.TextRange.Text
            $t = $t -replace "\r","`n"
            foreach ($ln in ($t -split "`n")) {
                if ($ln.Trim().Length -gt 0) { [void]$sb.AppendLine("$pad    | $ln") }
            }
        }
        if ($stype -eq 6) {
            foreach ($g in $shp.GroupItems) { Dump-Shape $g ($indent+1) }
        }
        # table
        try {
            if ($shp.HasTable) {
                $tb = $shp.Table
                for ($r=1; $r -le $tb.Rows.Count; $r++) {
                    $rowcells = @()
                    for ($c=1; $c -le $tb.Columns.Count; $c++) {
                        $ct = ""
                        try { $ct = $tb.Cell($r,$c).Shape.TextFrame.TextRange.Text } catch {}
                        $ct = ($ct -replace "[\r\n]"," ").Trim()
                        $rowcells += $ct
                    }
                    [void]$sb.AppendLine("$pad    TBL[$r] | " + ($rowcells -join " || "))
                }
            }
        } catch {}
    }

    $n = $pres.Slides.Count
    [void]$sb.AppendLine("TOTAL SLIDES: $n")
    [void]$sb.AppendLine("==================================================")
    for ($i=1; $i -le $n; $i++) {
        $slide = $pres.Slides.Item($i)
        [void]$sb.AppendLine("")
        [void]$sb.AppendLine("################## SLIDE $i ##################")
        foreach ($shp in $slide.Shapes) { Dump-Shape $shp 0 }
        # notes
        try {
            $notes = ""
            if ($slide.NotesPage.Shapes.Count -gt 0) {
                foreach ($ns in $slide.NotesPage.Shapes) {
                    if ($ns.HasTextFrame -and $ns.TextFrame.HasText) {
                        $nt = $ns.TextFrame.TextRange.Text
                        if ($nt.Trim().Length -gt 5 -and $ns.Type -ne 14) { $notes += $nt }
                    }
                }
            }
            if ($notes.Trim().Length -gt 0) {
                [void]$sb.AppendLine("  --- NOTES ---")
                [void]$sb.AppendLine("  " + ($notes -replace "\r","`n  "))
            }
        } catch {}
    }

    $out = "C:\Users\namma\.claude\plm_slide_work\v2_full_dump.txt"
    [System.IO.File]::WriteAllText($out, $sb.ToString(), [System.Text.Encoding]::UTF8)
    Write-Output "WROTE: $out"
    $pres.Close()
} finally {
    $ppt.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($ppt) | Out-Null
}
