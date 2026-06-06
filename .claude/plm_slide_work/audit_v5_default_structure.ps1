$ErrorActionPreference = 'Stop'
$ppt = 'D:\이력서\AVEVA - Marine Principal Technical Support & Consultant – PLM SME, Busan\Proposal\Future_Industrial_PLM_Meeting_Deck_EN.pptx'
$out = 'C:\Users\namma\.claude\plm_slide_work\v5_default_structure_audit.txt'
$pp = $null; $pres = $null
function ShapeText($sh) {
    $txt = ''
    try { if ($sh.HasTextFrame -and $sh.TextFrame.HasText) { $txt = $sh.TextFrame.TextRange.Text } } catch {}
    return (($txt -replace "[`r`n]+", ' ') -replace '\s+', ' ').Trim()
}
try {
    $pp = New-Object -ComObject PowerPoint.Application
    $pp.Visible = -1
    $pres = $pp.Presentations.Open($ppt, $false, $true, $false)
    $lines = New-Object System.Collections.Generic.List[string]
    $lines.Add("DECK=$ppt")
    $lines.Add("SLIDES=$($pres.Slides.Count)")
    for ($i = 1; $i -le $pres.Slides.Count; $i++) {
        $slide = $pres.Slides.Item($i)
        $texts = New-Object System.Collections.Generic.List[string]
        $pics = 0; $tables = 0; $textShapes = 0
        foreach ($sh in @($slide.Shapes)) {
            try { if ($sh.Type -eq 13) { $pics++ } } catch {}
            try { if ($sh.HasTable) { $tables++ } } catch {}
            $txt = ShapeText $sh
            if ($txt.Length -gt 0) { $textShapes++; $texts.Add($txt) }
        }
        $title = if ($texts.Count -gt 0) { $texts[0] } else { '<no text>' }
        if ($title.Length -gt 100) { $title = $title.Substring(0,100) }
        $lines.Add(('[{0:00}] pics={1} tables={2} textShapes={3} :: {4}' -f $i,$pics,$tables,$textShapes,$title))
        if ($i -ge 13 -and $i -le 15) {
            $lines.Add('  TEXT:')
            foreach ($t in $texts) {
                if ($t.Length -gt 220) { $lines.Add('    ' + $t.Substring(0,220)) } else { $lines.Add('    ' + $t) }
            }
        }
        if ($i -eq 2) {
            $lines.Add('  AGENDA/TEXT:')
            foreach ($t in $texts) { $lines.Add('    ' + $t) }
        }
    }
    [System.IO.File]::WriteAllLines($out, $lines, [System.Text.Encoding]::UTF8)
    Get-Content -LiteralPath $out
}
finally {
    if ($pres) { $pres.Close() }
    if ($pp) { $pp.Quit() }
}
