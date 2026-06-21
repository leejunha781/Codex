$ErrorActionPreference = "Stop"

$Path = "D:\이력서\ABB - Portfolio and Industry Manager\Joonha_Lee_ABB_Portfolio_Industry_Manager_Marine_Ports_Industrial_Automation_Resume.docx"
$OutPath = "C:\Users\namma\.claude\abb_resume_work\source_text.txt"

function New-WordApplication {
    try {
        return (New-Object -ComObject Word.Application)
    }
    catch {
        Start-Process -FilePath "C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE" -ArgumentList "/automation" -WindowStyle Hidden
        Start-Sleep -Seconds 5
        return (New-Object -ComObject Word.Application)
    }
}

$word = New-WordApplication
$word.Visible = $false
try { $word.DisplayAlerts = 0 } catch {}
try { $word.EventsEnabled = $false } catch {}
try { $word.AutomationSecurity = 3 } catch {}
$doc = $null

try {
    $confirmConversions = $false
    $readOnly = $true
    $addToRecentFiles = $false
    $doc = $word.Documents.Open([ref]$Path, [ref]$confirmConversions, [ref]$readOnly, [ref]$addToRecentFiles)
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
