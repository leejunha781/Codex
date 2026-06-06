$ErrorActionPreference = 'Stop'

$outPath = 'C:\Users\namma\.claude\plm_slide_work\active_v8_order_after_timeout.txt'

function Get-ShapeText {
    param([object]$Shape)
    $texts = New-Object System.Collections.Generic.List[string]
    try {
        if ($Shape.Type -eq 6) {
            for ($i = 1; $i -le $Shape.GroupItems.Count; $i++) {
                foreach ($t in (Get-ShapeText -Shape $Shape.GroupItems.Item($i))) { $texts.Add($t) }
            }
            return $texts
        }
    } catch {}
    try {
        if (($Shape.HasTextFrame -eq -1) -and ($Shape.TextFrame.HasText -eq -1)) {
            $txt = ($Shape.TextFrame.TextRange.Text -replace "[`r`n]+", " " -replace "\s+", " ").Trim()
            if ($txt.Length -gt 0) { $texts.Add($txt) }
        }
    } catch {}
    return $texts
}

$pp = [Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application')
if ($pp.Presentations.Count -lt 1) { throw 'No open presentations' }
$presentation = $pp.Presentations.Item(1)

$lines = New-Object System.Collections.Generic.List[string]
$lines.Add(("Presentation: {0}" -f $presentation.FullName))
$lines.Add(("SavedBefore: {0}" -f $presentation.Saved))
$lines.Add(("SlideCount: {0}" -f $presentation.Slides.Count))
$lines.Add('')
for ($i = 1; $i -le $presentation.Slides.Count; $i++) {
    $slide = $presentation.Slides.Item($i)
    $title = ''
    for ($s = 1; $s -le $slide.Shapes.Count; $s++) {
        foreach ($txt in (Get-ShapeText -Shape $slide.Shapes.Item($s))) {
            if ($txt -ne 'Future Industrial PLM' -and $txt -notmatch '^\d+$') {
                $title = $txt
                break
            }
        }
        if ($title.Length -gt 0) { break }
    }
    $lines.Add(("{0}. {1}" -f $i, $title))
}

$presentation.Save()
$lines.Insert(3, ("SavedAfter: {0}" -f $presentation.Saved))
[System.IO.File]::WriteAllLines($outPath, $lines, [System.Text.Encoding]::UTF8)
$presentation.Close()
if ($pp.Presentations.Count -eq 0) { $pp.Quit() }

Get-Content -LiteralPath $outPath
