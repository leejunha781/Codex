$ErrorActionPreference = "Stop"

$Path = "D:\이력서\에이치알브로스 - 무선통신 전문기업 통신장비 연구소장\이준하_통신장비_연구소장_사람인양식_이력서_자기소개서정정.docx"
$OutPath = "C:\Users\namma\.claude\hrbros_resume_work\source_text.txt"

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
    Write-Host "OPENING"
    $confirmConversions = $false
    $readOnly = $true
    $addToRecentFiles = $false
    $doc = $word.Documents.Open([ref]$Path, [ref]$confirmConversions, [ref]$readOnly, [ref]$addToRecentFiles)
    Write-Host "OPENED"
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
        if (($idx % 50) -eq 0) { Write-Host ("PARA " + $idx) }
        $text = ($p.Range.Text -replace "[\r\n\a\x07]", "" -replace "\s+", " ").Trim()
        if ($text.Length -gt 0) {
            $idx++
            $lines.Add(("P{0}: {1}" -f $idx, $text))
        }
    }

    $lines.Add("")
    $lines.Add("== TABLES ==")
    for ($ti = 1; $ti -le $doc.Tables.Count; $ti++) {
        Write-Host ("TABLE " + $ti)
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
