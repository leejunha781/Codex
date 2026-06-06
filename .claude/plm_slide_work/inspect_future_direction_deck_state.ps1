$ErrorActionPreference = 'Stop'

$proposal = 'D:\이력서\AVEVA - Marine Principal Technical Support & Consultant – PLM SME, Busan\Proposal'
$deckNames = @(
    'Future_Industrial_PLM_Meeting_Deck_EN_V6.pptx',
    'Future_Industrial_PLM_Meeting_Deck_EN.pptx',
    'Future_Industrial_PLM_Meeting_Deck_EN_Final.pptx',
    'Future_Industrial_PLM_Meeting_Deck_EN_V7.pptx'
)

$items = New-Object System.Collections.Generic.List[object]
foreach ($name in $deckNames) {
    $path = Join-Path $proposal $name
    if (Test-Path -LiteralPath $path) {
        $item = Get-Item -LiteralPath $path
        $items.Add([pscustomobject]@{
            Name = $item.Name
            Length = $item.Length
            LastWriteTime = $item.LastWriteTime
            FullName = $item.FullName
        })
    }
}

$target = Join-Path $proposal 'Future_Industrial_PLM_Meeting_Deck_EN_V6.pptx'
$pp = New-Object -ComObject PowerPoint.Application
$presentation = $null
try {
    $presentation = $pp.Presentations.Open($target, $true, $false, $false)
    $noteCount = 0
    $titles = New-Object System.Collections.Generic.List[string]
    for ($i = 1; $i -le $presentation.Slides.Count; $i++) {
        $slide = $presentation.Slides.Item($i)
        $title = ''
        for ($s = 1; $s -le $slide.Shapes.Count; $s++) {
            $shape = $slide.Shapes.Item($s)
            try {
                if (($shape.HasTextFrame -eq -1) -and ($shape.TextFrame.HasText -eq -1)) {
                    $txt = ($shape.TextFrame.TextRange.Text -replace "[`r`n]+", " " -replace "\s+", " ").Trim()
                    if ($txt.Length -gt 0 -and $txt -ne 'Future Industrial PLM' -and $txt -notmatch '^\d+$') {
                        $title = $txt
                        break
                    }
                }
            } catch {}
        }

        try {
            $notesText = $slide.NotesPage.Shapes.Placeholders(2).TextFrame.TextRange.Text
            if (($notesText -replace "\s+", '').Length -gt 0) {
                $noteCount++
            }
        } catch {}

        if ($i -le 5 -or $i -ge 33) {
            $titles.Add(("{0}. {1}" -f $i, $title))
        }
    }

    [pscustomobject]@{
        Target = $target
        SlideCount = $presentation.Slides.Count
        SlidesWithNotes = $noteCount
        SampleTitles = ($titles -join ' | ')
        Files = $items
    } | Format-List
} finally {
    if ($presentation -ne $null) { $presentation.Close() }
    if ($pp.Presentations.Count -eq 0) { $pp.Quit() }
}
