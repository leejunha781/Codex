$ErrorActionPreference = "Stop"

$Path = "D:\이력서\AVEVA - Business Development Manager\Joonha_Lee_AVEVA_BDM_Engineering_Marine_Korea_Japan_Seoul_Resume_Formal_Tone.docx"
$OutPath = "C:\Users\namma\.claude\aveva_bdm_resume_work\source_resume_text.txt"

$word = New-Object -ComObject Word.Application
$word.Visible = $false
$doc = $null

try {
    $doc = $word.Documents.Open($Path, $false, $true)
    $lines = New-Object System.Collections.Generic.List[string]
    $lines.Add("PATH: $Path")
    $lines.Add("PAGES: " + $doc.ComputeStatistics(2))
    $lines.Add("WORDS: " + $doc.ComputeStatistics(0))
    $lines.Add("TABLES: " + $doc.Tables.Count)
    $lines.Add("PARAGRAPHS: " + $doc.Paragraphs.Count)
    $lines.Add("")
    $lines.Add("== PARAGRAPHS ==")

    $idx = 0
    foreach ($p in $doc.Paragraphs) {
        $text = ($p.Range.Text -replace "[\r\n\a\x07]", "" -replace "\s+", " ").Trim()
        if ($text.Length -gt 0) {
            $idx++
            $lines.Add(("P{0}: {1}" -f $idx, $text))
        }
    }

    $lines.Add("")
    $lines.Add("== TABLES ==")
    for ($ti = 1; $ti -le $doc.Tables.Count; $ti++) {
        $table = $doc.Tables.Item($ti)
        $lines.Add(("TABLE {0}: rows={1}, cols={2}" -f $ti, $table.Rows.Count, $table.Columns.Count))
        $ci = 0
        foreach ($cell in $table.Range.Cells) {
            $ci++
            $text = ($cell.Range.Text -replace "[\r\n\a\x07]", "" -replace "\s+", " ").Trim()
            if ($text.Length -gt 0) {
                $lines.Add(("  C{0}: {1}" -f $ci, $text))
            }
        }
    }

    [System.IO.File]::WriteAllLines($OutPath, $lines, [System.Text.Encoding]::UTF8)
    Write-Host "Extracted to $OutPath"
}
finally {
    if ($doc -ne $null) {
        try { $doc.Close([ref]$false) } catch {}
    }
    if ($word -ne $null -and $word.Documents.Count -eq 0) {
        $word.Quit()
    }
}
