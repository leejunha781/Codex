# Dump AVEVA BDM resume + codex input docx text, and V18 deck text
$ErrorActionPreference = 'Stop'
$out = 'C:\Users\namma\.claude\aveva_bdm_work'
if (-not (Test-Path $out)) { New-Item -ItemType Directory -Path $out | Out-Null }

$docs = @(
    @{ Path = 'D:\이력서\AVEVA - Business Development Manager\Joonha_Lee_AVEVA_BDM_Engineering_Marine_Korea_Japan_Seoul_Resume_Formal_Tone.docx'; Out = "$out\resume_formal_tone.txt" },
    @{ Path = 'D:\이력서\AVEVA - Business Development Manager\AVEVA_BDM_Resume_Codex_Input.docx'; Out = "$out\codex_input.txt" },
    @{ Path = 'D:\이력서\AVEVA - Business Development Manager\Joonha_Lee_AVEVA_BDM_Engineering_Marine_Korea_Japan_Seoul_Resume.docx'; Out = "$out\resume_plain.txt" }
)

$word = New-Object -ComObject Word.Application
$word.Visible = $false
try {
    foreach ($d in $docs) {
        $doc = $word.Documents.Open($d.Path, $false, $true)  # ReadOnly
        try {
            $sb = New-Object System.Text.StringBuilder
            [void]$sb.AppendLine("=== FILE: $($d.Path)")
            [void]$sb.AppendLine("=== Paragraph count: $($doc.Paragraphs.Count); Tables: $($doc.Tables.Count)")
            $i = 0
            foreach ($p in $doc.Paragraphs) {
                $i++
                $txt = ($p.Range.Text -replace "[\r\n\a\x0b\x07]", '').TrimEnd()
                $style = ''
                try { $style = $p.Style.NameLocal } catch {}
                $inTable = $false
                try { $inTable = $p.Range.Information(12) } catch {}  # wdWithInTable
                $tag = if ($inTable) { 'T' } else { ' ' }
                [void]$sb.AppendLine(("[{0:d3}|{1}|{2}] {3}" -f $i, $tag, $style, $txt))
            }
            [System.IO.File]::WriteAllText($d.Out, $sb.ToString(), [System.Text.Encoding]::UTF8)
            Write-Output "DUMPED: $($d.Out) ($i paragraphs)"
        } finally {
            $doc.Close($false)
        }
    }
} finally {
    if ($word.Documents.Count -eq 0) { $word.Quit() }
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($word)
}

# --- PPTX V18 dump ---
$pptPath = 'D:\이력서\AVEVA - Marine Principal Technical Support & Consultant – PLM SME, Busan\Proposal\Future_Industrial_PLM_Meeting_Deck_EN_V18.pptx'
$ppt = New-Object -ComObject PowerPoint.Application
try {
    $pres = $ppt.Presentations.Open($pptPath, $true, $false, $false)  # ReadOnly, untitled=no, window=no
    try {
        $sb2 = New-Object System.Text.StringBuilder
        [void]$sb2.AppendLine("=== DECK: $pptPath — $($pres.Slides.Count) slides")
        foreach ($slide in $pres.Slides) {
            [void]$sb2.AppendLine("")
            [void]$sb2.AppendLine("########## SLIDE $($slide.SlideIndex) ##########")
            foreach ($shape in $slide.Shapes) {
                if ($shape.HasTextFrame -eq -1 -and $shape.TextFrame.HasText -eq -1) {
                    $t = $shape.TextFrame.TextRange.Text -replace "[\r\n\x0b]+", ' | '
                    [void]$sb2.AppendLine("  [$($shape.Name)] $t")
                }
                if ($shape.HasTable -eq -1) {
                    $tb = $shape.Table
                    for ($r = 1; $r -le $tb.Rows.Count; $r++) {
                        $cells = @()
                        for ($c = 1; $c -le $tb.Columns.Count; $c++) {
                            try { $cells += $tb.Cell($r, $c).Shape.TextFrame.TextRange.Text -replace "[\r\n\x0b]+", ' ' } catch {}
                        }
                        [void]$sb2.AppendLine("  [TABLE r$r] " + ($cells -join ' || '))
                    }
                }
                if ($shape.Type -eq 6) {  # group
                    foreach ($gs in $shape.GroupItems) {
                        if ($gs.HasTextFrame -eq -1 -and $gs.TextFrame.HasText -eq -1) {
                            $t = $gs.TextFrame.TextRange.Text -replace "[\r\n\x0b]+", ' | '
                            [void]$sb2.AppendLine("  [GRP:$($gs.Name)] $t")
                        }
                    }
                }
            }
        }
        [System.IO.File]::WriteAllText("$out\deck_v18.txt", $sb2.ToString(), [System.Text.Encoding]::UTF8)
        Write-Output "DUMPED: $out\deck_v18.txt ($($pres.Slides.Count) slides)"
    } finally {
        $pres.Saved = $true
        $pres.Close()
    }
} finally {
    if ($ppt.Presentations.Count -eq 0) { $ppt.Quit() }
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($ppt)
}
Write-Output 'ALL DONE'
