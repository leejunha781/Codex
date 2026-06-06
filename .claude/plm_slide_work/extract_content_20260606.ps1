$ErrorActionPreference = "Stop"
$path = "D:\이력서\AVEVA - Marine Principal Technical Support & Consultant – PLM SME, Busan\Proposal\Future_Industrial_PLM_Meeting_Deck_EN_Rebuilt_v2.pptx"
$out  = "C:\Users\namma\.claude\plm_slide_work\deck_content_20260606.txt"

$ppt = New-Object -ComObject PowerPoint.Application
$sb = New-Object System.Text.StringBuilder

function Get-ShapeText($shape, $depth) {
    $pad = "  " * $depth
    try {
        if ($shape.Type -eq 6) { # msoGroup
            foreach ($s in $shape.GroupItems) { Get-ShapeText $s ($depth+1) }
            return
        }
        if ($shape.HasTable) {
            $tb = $shape.Table
            for ($r=1; $r -le $tb.Rows.Count; $r++) {
                $rowtxt = @()
                for ($c=1; $c -le $tb.Columns.Count; $c++) {
                    try { $t = $tb.Cell($r,$c).Shape.TextFrame.TextRange.Text } catch { $t = "" }
                    $rowtxt += ($t -replace "[\r\n\v]"," ")
                }
                [void]$sb.AppendLine($pad + "| " + ($rowtxt -join " | "))
            }
            return
        }
        if ($shape.HasTextFrame -and $shape.TextFrame.HasText) {
            $txt = $shape.TextFrame.TextRange.Text
            foreach ($line in ($txt -split "[\r\n\v]")) {
                $l = $line.Trim()
                if ($l.Length -gt 0) { [void]$sb.AppendLine($pad + $l) }
            }
        }
    } catch {}
}

try {
    $doc = $ppt.Presentations.Open($path, $true, $false, $false) # ReadOnly, Untitled=F, WithWindow=F
    $n = $doc.Slides.Count
    [void]$sb.AppendLine("DECK: $path")
    [void]$sb.AppendLine("TOTAL SLIDES: $n")
    [void]$sb.AppendLine("")
    foreach ($slide in $doc.Slides) {
        $idx = $slide.SlideIndex
        [void]$sb.AppendLine("============================================================")
        [void]$sb.AppendLine("SLIDE $idx")
        [void]$sb.AppendLine("============================================================")
        # order shapes top-to-bottom, left-to-right for readability
        $shapes = @($slide.Shapes) | Sort-Object { [int]$_.Top }, { [int]$_.Left }
        foreach ($shape in $shapes) { Get-ShapeText $shape 0 }
        # speaker notes
        try {
            $notes = $slide.NotesPage.Shapes.Placeholders.Item(2).TextFrame.TextRange.Text
            if ($notes -and $notes.Trim().Length -gt 0) {
                [void]$sb.AppendLine("  [NOTES] " + ($notes -replace "[\r\n\v]"," "))
            }
        } catch {}
        [void]$sb.AppendLine("")
    }
    $doc.Close()
} finally {
    $ppt.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ppt) | Out-Null
}
[System.IO.File]::WriteAllText($out, $sb.ToString(), (New-Object System.Text.UTF8Encoding($false)))
Write-Output "WROTE: $out"
Write-Output ("CHARS: " + $sb.Length)
