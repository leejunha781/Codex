$ErrorActionPreference = 'Stop'

$deckPath = 'D:\이력서\AVEVA - Marine Principal Technical Support & Consultant – PLM SME, Busan\Proposal\Future_Industrial_PLM_Meeting_Deck_EN_V6.pptx'
$outPath = 'C:\Users\namma\.claude\plm_slide_work\v6_future_direction_outline.txt'

function Add-ShapeText {
    param(
        [object]$Shape,
        [System.Collections.Generic.List[string]]$Lines,
        [string]$Prefix
    )

    try {
        if ($Shape.Type -eq 6) {
            for ($i = 1; $i -le $Shape.GroupItems.Count; $i++) {
                Add-ShapeText -Shape $Shape.GroupItems.Item($i) -Lines $Lines -Prefix $Prefix
            }
            return
        }
    } catch {}

    try {
        if ($Shape.HasTable -eq -1) {
            $tbl = $Shape.Table
            for ($r = 1; $r -le $tbl.Rows.Count; $r++) {
                $cells = New-Object System.Collections.Generic.List[string]
                for ($c = 1; $c -le $tbl.Columns.Count; $c++) {
                    try {
                        $txt = $tbl.Cell($r, $c).Shape.TextFrame.TextRange.Text
                        $txt = ($txt -replace "[`r`n]+", " " -replace "\s+", " ").Trim()
                        if ($txt.Length -gt 0) { $cells.Add($txt) }
                    } catch {}
                }
                if ($cells.Count -gt 0) {
                    $Lines.Add($Prefix + ($cells -join ' | '))
                }
            }
        }
    } catch {}

    try {
        if (($Shape.HasTextFrame -eq -1) -and ($Shape.TextFrame.HasText -eq -1)) {
            $txt = $Shape.TextFrame.TextRange.Text
            $txt = ($txt -replace "[`r`n]+", " " -replace "\s+", " ").Trim()
            if ($txt.Length -gt 0) {
                $Lines.Add($Prefix + $txt)
            }
        }
    } catch {}
}

$pp = New-Object -ComObject PowerPoint.Application
$presentation = $null
try {
    $presentation = $pp.Presentations.Open($deckPath, $true, $false, $false)
    $lines = New-Object System.Collections.Generic.List[string]
    $lines.Add("Deck: $deckPath")
    $lines.Add("SlideCount: $($presentation.Slides.Count)")
    $lines.Add(("Generated: {0:yyyy-MM-dd HH:mm:ss}" -f (Get-Date)))
    $lines.Add('')

    for ($s = 1; $s -le $presentation.Slides.Count; $s++) {
        $slide = $presentation.Slides.Item($s)
        $lines.Add("===== SLIDE $s =====")
        for ($i = 1; $i -le $slide.Shapes.Count; $i++) {
            Add-ShapeText -Shape $slide.Shapes.Item($i) -Lines $lines -Prefix ''
        }
        $lines.Add('')
    }

    [System.IO.File]::WriteAllLines($outPath, $lines, [System.Text.Encoding]::UTF8)
    Write-Output $outPath
} finally {
    if ($presentation -ne $null) { $presentation.Close() }
    if ($pp.Presentations.Count -eq 0) { $pp.Quit() }
}
