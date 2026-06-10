# Insert formula blocks (from formula_spec.json) into the interview-prep docx -> _v2
$ErrorActionPreference = "Stop"
$src  = "D:\이력서\모아컴코리아 - HW 개발\이준하_모아컴코리아_HW설계_전문면접_기술이론_전략_수식포함.docx"
$dst  = "D:\이력서\모아컴코리아 - HW 개발\이준하_모아컴코리아_HW설계_전문면접_기술이론_전략_수식포함_v2.docx"
$pdf  = "D:\이력서\모아컴코리아 - HW 개발\이준하_모아컴코리아_HW설계_전문면접_기술이론_전략_수식포함_v2.pdf"
$spec = Get-Content "C:\Users\namma\.claude\moacom_work\formula_spec.json" -Raw -Encoding UTF8 | ConvertFrom-Json

function NormText([string]$t) { ($t -replace "[\r\n\a\x07]","" -replace "\s+"," ").Trim() }

$word = New-Object -ComObject Word.Application
$word.Visible = $false
try {
    $doc = $word.Documents.Open($src, $false, $true)
    $doc.SaveAs2($dst)
    Write-Output "SavedAs: $dst"

    # pass 0: in-place partial edits (e.g. scan loss sign fix)
    foreach ($e in $spec.edits) {
        $n = 0
        foreach ($p in $doc.Paragraphs) {
            $norm = NormText $p.Range.Text
            if (-not $norm.Contains($e.anchor)) { continue }
            $r = $p.Range.Duplicate
            $r.End = $r.End - 1
            $cur = $r.Text
            if ($cur.Contains($e.from)) {
                $r.Text = $cur.Replace($e.from, $e.to)
                $n++
            }
        }
        Write-Output ("EDIT {0}: {1} occurrence(s)" -f $e.id, $n)
    }

    # pass 1: locate every anchor (must be unique)
    $targets = @()
    foreach ($ins in $spec.inserts) { $ins | Add-Member -NotePropertyName Hits -NotePropertyValue 0 -Force }
    foreach ($p in $doc.Paragraphs) {
        $norm = NormText $p.Range.Text
        if ($norm.Length -eq 0) { continue }
        foreach ($ins in $spec.inserts) {
            $isHit = if ($ins.PSObject.Properties["anchorExact"] -and $ins.anchorExact) { $norm -eq $ins.anchor } else { $norm.Contains($ins.anchor) }
            if ($isHit) {
                $ins.Hits++
                $targets += [pscustomobject]@{ Id = $ins.id; Start = $p.Range.Start; Lines = $ins.lines }
            }
        }
    }
    $bad = $false
    foreach ($ins in $spec.inserts) {
        if ($ins.Hits -ne 1) { Write-Output ("ANCHOR PROBLEM: {0} hits={1}" -f $ins.id, $ins.Hits); $bad = $true }
    }
    if ($bad) { throw "anchor resolution failed - aborting before any insert" }

    # pass 2: insert in reverse document order so earlier Starts stay valid
    foreach ($t in ($targets | Sort-Object Start -Descending)) {
        $clean = @()
        foreach ($l in $t.Lines) { $clean += ($l -replace "^## ", "") }
        $r = $doc.Range($t.Start, $t.Start)
        $r.InsertBefore((($clean -join "`r") + "`r"))
        Write-Output ("INSERTED {0} ({1} line(s))" -f $t.Id, $clean.Count)
    }

    # pass 3: restyle inserted paragraphs (match by normalized text)
    $bodySet = @{}; $subSet = @{}
    foreach ($ins in $spec.inserts) {
        foreach ($l in $ins.lines) {
            if ($l -match "^## ") { $subSet[(NormText ($l -replace "^## ", ""))] = $true }
            else { $bodySet[(NormText $l)] = $true }
        }
    }
    $styledBody = 0; $styledSub = 0
    foreach ($p in $doc.Paragraphs) {
        $norm = NormText $p.Range.Text
        if ($norm.Length -eq 0) { continue }
        $r2 = $null
        if ($bodySet.ContainsKey($norm)) {
            $r2 = $p.Range.Duplicate; $r2.End = $r2.End - 1
            $r2.Font.Name = "맑은 고딕"; $r2.Font.Size = [single]10; $r2.Font.Bold = [int]0
            $p.Format.LineSpacingRule = 0
            $styledBody++
        } elseif ($subSet.ContainsKey($norm)) {
            $r2 = $p.Range.Duplicate; $r2.End = $r2.End - 1
            $r2.Font.Name = "맑은 고딕"; $r2.Font.Size = [single]11; $r2.Font.Bold = [int]1
            $p.Format.LineSpacingRule = 0
            $styledSub++
        }
    }
    $expBody = 0; $expSub = 0
    foreach ($ins in $spec.inserts) {
        foreach ($l in $ins.lines) { if ($l -match "^## ") { $expSub++ } else { $expBody++ } }
    }
    Write-Output ("STYLED body {0}/{1}  sub {2}/{3}" -f $styledBody, $expBody, $styledSub, $expSub)

    # heading survival check
    $headings = @()
    foreach ($p in $doc.Paragraphs) {
        $sz = $p.Range.Font.Size
        if ($sz -eq 14 -or $sz -eq 18) {
            $h = NormText $p.Range.Text
            if ($h.Length -gt 0) { $headings += $h }
        }
    }
    Write-Output ("HeadingCount: {0}" -f $headings.Count)
    Write-Output ($headings -join " | ")

    $doc.Save()
    Write-Output ("PAGES: {0}  WORDS: {1}" -f $doc.ComputeStatistics(2), $doc.ComputeStatistics(0))
    $doc.ExportAsFixedFormat($pdf, 17)
    Write-Output "PDF: $pdf"
    $doc.Close([ref]$false)
} finally {
    if ($word.Documents.Count -eq 0) { $word.Quit() }
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
}
