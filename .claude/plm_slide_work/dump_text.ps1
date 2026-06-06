$ErrorActionPreference = "Stop"
$path = "D:\이력서\AVEVA - Marine Principal Technical Support & Consultant – PLM SME, Busan\Proposal\Future_Industrial_PLM_Meeting_Deck_EN_Rebuilt_v2.pptx"
$out  = "C:\Users\namma\.claude\plm_slide_work\deck_dump.txt"

$ppt = New-Object -ComObject PowerPoint.Application
$pres = $ppt.Presentations.Open($path, $true, $false, $false)  # ReadOnly=T, Untitled=F, WithWindow=F

$sb = New-Object System.Text.StringBuilder
[void]$sb.AppendLine("TOTAL SLIDES: " + $pres.Slides.Count)
[void]$sb.AppendLine("SLIDE SIZE: " + $pres.PageSetup.SlideWidth + " x " + $pres.PageSetup.SlideHeight)
[void]$sb.AppendLine("")

function Get-ShapeText($shp, $sb, $indent) {
    $pre = "  " * $indent
    try {
        if ($shp.Type -eq 6) {  # msoGroup
            [void]$sb.AppendLine($pre + "[GROUP: " + $shp.Name + "]")
            foreach ($g in $shp.GroupItems) { Get-ShapeText $g $sb ($indent+1) }
            return
        }
    } catch {}
    # Table?
    try {
        if ($shp.HasTable -eq -1) {
            $tb = $shp.Table
            [void]$sb.AppendLine($pre + "[TABLE " + $tb.Rows.Count + "x" + $tb.Columns.Count + ": " + $shp.Name + "]")
            for ($r=1; $r -le $tb.Rows.Count; $r++) {
                $rowtxt = @()
                for ($c=1; $c -le $tb.Columns.Count; $c++) {
                    $t = ""
                    try { $t = $tb.Cell($r,$c).Shape.TextFrame.TextRange.Text } catch {}
                    $rowtxt += ($t -replace "[\r\n\v]"," / ")
                }
                [void]$sb.AppendLine($pre + "  R$r | " + ($rowtxt -join " || "))
            }
            return
        }
    } catch {}
    # Text frame
    try {
        if ($shp.HasTextFrame -eq -1 -and $shp.TextFrame.HasText -eq -1) {
            $txt = $shp.TextFrame.TextRange.Text
            $txt = ($txt -replace "\r","`n")
            $lines = $txt -split "`n"
            $clean = ($lines | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne "" }) -join " | "
            [void]$sb.AppendLine($pre + "<" + $shp.Name + "> " + $clean)
            return
        }
    } catch {}
    # Picture / other
    try {
        $tn = ""
        switch ($shp.Type) { 13 {$tn="PICTURE"} 17 {$tn="TEXTBOX-empty"} default {$tn="shape-type-"+$shp.Type} }
        if ($shp.Type -eq 13) { [void]$sb.AppendLine($pre + "{PICTURE: " + $shp.Name + " w=" + [math]::Round($shp.Width) + " h=" + [math]::Round($shp.Height) + " l=" + [math]::Round($shp.Left) + " t=" + [math]::Round($shp.Top) + "}") }
    } catch {}
}

for ($i=1; $i -le $pres.Slides.Count; $i++) {
    $slide = $pres.Slides.Item($i)
    [void]$sb.AppendLine("============================================================")
    [void]$sb.AppendLine("SLIDE $i  (layout: " + $slide.Layout + ", shapes: " + $slide.Shapes.Count + ")")
    [void]$sb.AppendLine("============================================================")
    foreach ($shp in $slide.Shapes) { Get-ShapeText $shp $sb 0 }
    # notes
    try {
        $nt = $slide.NotesPage.Shapes.Placeholders.Item(2).TextFrame.TextRange.Text
        if ($nt.Trim() -ne "") { [void]$sb.AppendLine("  ((NOTES)): " + ($nt -replace "[\r\n]"," ")) }
    } catch {}
    [void]$sb.AppendLine("")
}

[System.IO.File]::WriteAllText($out, $sb.ToString(), [System.Text.Encoding]::UTF8)
$pres.Close()
if ($ppt.Presentations.Count -eq 0) { } # never Quit
Write-Output "DONE -> $out"
