# Office COM Patterns

Use these patterns when working on resumes, proposals, and application decks in the Windows Office COM environment.

## Environment

- Use Windows PowerShell 5.1 syntax. If Codex is running in PowerShell 7, call `powershell.exe -NoProfile -Command ...` for COM and WinRT PDF rendering scripts.
- Do not use Python, Node.js, LibreOffice, or shell chains such as `&&`.
- Run `.ps1` scripts without execution-policy changes:

```powershell
$scriptBlock = [scriptblock]::Create((Get-Content -LiteralPath .\script.ps1 -Raw))
& $scriptBlock -InputPath "D:\path\file.docx"
```

## Word Session Pattern

```powershell
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$doc = $null
try {
    $doc = $word.Documents.Open($path, $false, $false)
    # edits
    $doc.Save()
    $doc.ExportAsFixedFormat($pdfPath, 17)
    $doc.Close([ref]$true)
} finally {
    if ($doc -ne $null) {
        try { $doc.Close([ref]$false) } catch {}
    }
    if ($word -ne $null -and $word.Documents.Count -eq 0) {
        $word.Quit()
    }
}
```

## Word Text Matching

Avoid `Find.Execute`; it can throw `DISP_E_TYPEMISMATCH` in PowerShell. Iterate paragraphs instead:

```powershell
foreach ($p in $doc.Paragraphs) {
    $text = ($p.Range.Text -replace "[\r\n\a\x07]", "" -replace "\s+", " ").Trim()
    if ($text -eq "Target heading") {
        # use $p.Range
    }
}
```

## Word Paragraph Replacement

Keep the paragraph mark to avoid merging with the next paragraph:

```powershell
$r = $p.Range.Duplicate
$r.End = $r.End - 1
$r.Text = $newText
$p.Range.Font.Bold = [int]0
```

If the lead-in should be bold, apply it explicitly after replacement:

```powershell
$lead = "Core Skills:"
$leadRange = $p.Range.Duplicate
$leadRange.End = $leadRange.Start + $lead.Length
$leadRange.Font.Bold = [int]1
```

## Word Tables

Iterate merged tables safely:

```powershell
foreach ($cell in $table.Range.Cells) {
    $cellText = ($cell.Range.Text -replace "[\r\n\a\x07]", "" -replace "\s+", " ").Trim()
}
```

After deleting a table, re-read nearby headings because `Tables.Delete()` can silently truncate adjacent heading text.

## Moving Word Blocks

Capture start/end positions, cut the range, re-find the anchor after cutting, then paste:

```powershell
$sourceStart = $sourceHeading.Range.Start
$sourceEnd = $nextHeading.Range.Start
$doc.Range($sourceStart, $sourceEnd).Cut()
# Re-find target anchor because positions shifted.
$target = $null
foreach ($p in $doc.Paragraphs) {
    $text = ($p.Range.Text -replace "[\r\n\a\x07]", "" -replace "\s+", " ").Trim()
    if ($text -eq $targetHeadingText) { $target = $p; break }
}
$pasteRange = $doc.Range($target.Range.Start, $target.Range.Start)
$pasteRange.Collapse(1)
$pasteRange.Paste()
```

## PowerPoint Export Pattern

```powershell
$ppt = New-Object -ComObject PowerPoint.Application
$presentation = $null
try {
    $presentation = $ppt.Presentations.Open($path, $true, $false, $false)
    $presentation.SaveAs($pdfPath, 32)
    $presentation.Close()
} finally {
    if ($presentation -ne $null) {
        try { $presentation.Close() } catch {}
    }
    if ($ppt -ne $null) {
        $ppt.Quit()
    }
}
```

## PDF Text Check With Word

```powershell
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$doc = $null
try {
    $doc = $word.Documents.Open($pdfPath, $false, $true)
    $text = $doc.Content.Text
    $doc.Close([ref]$false)
} finally {
    if ($doc -ne $null) {
        try { $doc.Close([ref]$false) } catch {}
    }
    if ($word -ne $null -and $word.Documents.Count -eq 0) {
        $word.Quit()
    }
}
```

## Orphan Cleanup

Only kill hidden, titleless Office processes when a COM run failed and left an orphan:

```powershell
Get-Process WINWORD | Where-Object { [string]::IsNullOrEmpty($_.MainWindowTitle) } | Stop-Process -Force
Get-Process POWERPNT | Where-Object { [string]::IsNullOrEmpty($_.MainWindowTitle) } | Stop-Process -Force
```
