$ErrorActionPreference = "Stop"
$path = "D:\이력서\AVEVA - Marine Principal Technical Support & Consultant – PLM SME, Busan\Proposal\Future_Industrial_PLM_Meeting_Deck_EN_V15.pptx"
$out  = "C:\Users\namma\.claude\plm_slide_work\dump_v15.txt"

$pp = New-Object -ComObject PowerPoint.Application
$pres = $pp.Presentations.Open($path, $true, $false, $false)  # ReadOnly, Untitled=false, WithWindow=false

$sb = New-Object System.Text.StringBuilder
[void]$sb.AppendLine("SLIDE COUNT: " + $pres.Slides.Count)

function Dump-Shape($shp, $sb, $indent) {
    $pre = "  " * $indent
    $nm = $shp.Name
    $hasTextLabel = ""
    try {
        if ($shp.HasTextFrame) {
            if ($shp.TextFrame.HasText) {
                $t = $shp.TextFrame.TextRange.Text
                $t = ($t -replace "[\r\v\x0b]", " / ") -replace "\s+", " "
                $t = $t.Trim()
                $hasTextLabel = $t
            }
        }
    } catch {}
    [void]$sb.AppendLine($pre + "[" + $nm + "] " + $hasTextLabel)
    # Group
    try {
        if ($shp.Type -eq 6) {  # msoGroup
            foreach ($g in $shp.GroupItems) { Dump-Shape $g $sb ($indent+1) }
        }
    } catch {}
    # Table
    try {
        if ($shp.HasTable) {
            $tb = $shp.Table
            for ($r=1; $r -le $tb.Rows.Count; $r++) {
                $rowtxt = @()
                for ($c=1; $c -le $tb.Columns.Count; $c++) {
                    $ct = ""
                    try { $ct = $tb.Cell($r,$c).Shape.TextFrame.TextRange.Text } catch {}
                    $ct = ($ct -replace "[\r\v\x0b]"," ") -replace "\s+"," "
                    $rowtxt += $ct.Trim()
                }
                [void]$sb.AppendLine($pre + "  TBL r$r | " + ($rowtxt -join " || "))
            }
        }
    } catch {}
}

for ($i=1; $i -le $pres.Slides.Count; $i++) {
    $sl = $pres.Slides.Item($i)
    [void]$sb.AppendLine("")
    [void]$sb.AppendLine("==================== SLIDE $i ====================")
    foreach ($shp in $sl.Shapes) { Dump-Shape $shp $sb 0 }
    # notes
    try {
        $notes = ""
        foreach ($ns in $sl.NotesPage.Shapes) {
            if ($ns.HasTextFrame -and $ns.TextFrame.HasText) {
                $nt = $ns.TextFrame.TextRange.Text
                if ($nt.Trim().Length -gt 0 -and $nt.Trim() -ne $i.ToString()) { $notes += $nt + " " }
            }
        }
        $notes = ($notes -replace "[\r\v\x0b]"," ") -replace "\s+"," "
        if ($notes.Trim().Length -gt 0) { [void]$sb.AppendLine("NOTES: " + $notes.Trim()) }
    } catch {}
}

$pres.Close()
$pp.Quit()
[System.IO.File]::WriteAllText($out, $sb.ToString(), [System.Text.Encoding]::UTF8)
Write-Output "DONE -> $out"
Write-Output ("Length chars: " + $sb.Length)
