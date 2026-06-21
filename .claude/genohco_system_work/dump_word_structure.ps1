param(
    [Parameter(Mandatory=$true)][string]$InputPath,
    [Parameter(Mandatory=$true)][string]$OutPath
)

$ErrorActionPreference = "Stop"

function Normalize-WordText {
    param([string]$Text)
    if ($null -eq $Text) { return "" }
    return (($Text -replace "[`r`n\a\x07]", "") -replace "\s+", " ").Trim()
}

$word = $null
for ($attempt = 1; $attempt -le 3 -and $word -eq $null; $attempt++) {
    try {
        $word = New-Object -ComObject Word.Application
    }
    catch {
        if ($attempt -eq 3) { throw }
        Start-Sleep -Seconds 5
    }
}
$word.Visible = $false
try { $word.DisplayAlerts = [int]0 } catch {}
$doc = $null
try {
    $doc = $word.Documents.Open($InputPath, $false, $true)
    $lines = New-Object System.Collections.Generic.List[string]
    $lines.Add("FILE: $InputPath")
    $lines.Add("FULLNAME: $($doc.FullName)")
    $lines.Add("PAGE_COUNT: $($doc.ComputeStatistics(2))")
    $lines.Add("WORD_COUNT: $($doc.ComputeStatistics(0))")
    $lines.Add("PARAGRAPHS: $($doc.Paragraphs.Count)")
    $lines.Add("TABLES: $($doc.Tables.Count)")
    $lines.Add("")
    $lines.Add("== PARAGRAPHS ==")
    for ($i = 1; $i -le $doc.Paragraphs.Count; $i++) {
        $p = $doc.Paragraphs.Item($i)
        $text = Normalize-WordText $p.Range.Text
        if ($text.Length -gt 0) {
            $styleName = ""
            try { $styleName = $p.Range.Style.NameLocal } catch {}
            $fontName = ""
            try { $fontName = $p.Range.Font.Name } catch {}
            $fontSize = ""
            try { $fontSize = [string]$p.Range.Font.Size } catch {}
            $bold = ""
            try { $bold = [string]$p.Range.Font.Bold } catch {}
            $lines.Add(("[P{0:0000}] style='{1}' font='{2}' size='{3}' bold='{4}' text='{5}'" -f $i, $styleName, $fontName, $fontSize, $bold, $text))
        }
    }
    $lines.Add("")
    $lines.Add("== TABLES ==")
    for ($t = 1; $t -le $doc.Tables.Count; $t++) {
        $tb = $doc.Tables.Item($t)
        $lines.Add(("[TABLE {0}] rows={1} cols={2} cells={3}" -f $t, $tb.Rows.Count, $tb.Columns.Count, $tb.Range.Cells.Count))
        $cIndex = 0
        foreach ($cell in $tb.Range.Cells) {
            $cIndex++
            $txt = Normalize-WordText $cell.Range.Text
            if ($txt.Length -gt 0) {
                $lines.Add(("  [T{0}.C{1:000}] '{2}'" -f $t, $cIndex, $txt))
            }
        }
    }
    [System.IO.File]::WriteAllLines($OutPath, $lines, [System.Text.Encoding]::UTF8)
    $doc.Close([ref]$false)
    $doc = $null
}
finally {
    if ($doc -ne $null) {
        try { $doc.Close([ref]$false) } catch {}
    }
    if ($word -ne $null) {
        try {
            if ($word.Documents.Count -eq 0) {
                $word.Quit()
            }
        } catch {}
    }
}
