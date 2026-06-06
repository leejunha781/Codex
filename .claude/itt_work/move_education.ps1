$ErrorActionPreference = "Stop"
$path = "D:\이력서\ITT Cannon\Joonha_Lee_ITT_Cannon_FAE_Resume_Final_Integrated_v2.docx"
$pdf  = "C:\Users\namma\.claude\itt_work\v2_current.pdf"
$log  = "C:\Users\namma\.claude\itt_work\move_education_log.txt"
$L = @()
$wdCollapseStart = 1
function CleanText($p) { return ($p.Range.Text -replace "[\r\n\a\x07]","" -replace "\s+"," ").Trim() }

$word = New-Object -ComObject Word.Application
$word.Visible = $false
try {
    $doc = $word.Documents.Open($path, $false, $false)

    # 1) Locate Education-heading start and Career-Summary-heading start (block = Education heading + its table + trailing empty)
    $eduStart = $null; $careerStart = $null
    foreach ($p in $doc.Paragraphs) {
        $t = CleanText $p
        if ($t.StartsWith("Education, Certifications, Language")) { $eduStart = $p.Range.Start }
        if ($t -eq "Career Summary" -and $careerStart -eq $null) { $careerStart = $p.Range.Start }
    }
    if ($eduStart -eq $null -or $careerStart -eq $null) { throw "edu/career bounds not found (edu=$eduStart career=$careerStart)" }
    if ($eduStart -ge $careerStart) { throw "unexpected order edu=$eduStart career=$careerStart" }
    $L += "Education block range = $eduStart .. $careerStart"

    # 2) Cut the Education block to clipboard
    $moveRange = $doc.Range($eduStart, $careerStart)
    $moveRange.Cut()
    $L += "Cut Education block."

    # 3) Find Compensation heading start (positions shifted after cut) and paste before it
    $compStart = $null
    foreach ($p in $doc.Paragraphs) {
        if ((CleanText $p).StartsWith("Compensation / Salary Information")) { $compStart = $p.Range.Start; break }
    }
    if ($compStart -eq $null) { throw "Compensation heading not found after cut" }
    $ins = $doc.Range($compStart, $compStart)
    $ins.Collapse($wdCollapseStart)
    $ins.Paste()
    $L += "Pasted Education block before Compensation heading."

    # ---- verify section heading order ----
    $names = @("Professional Profile","Key Strengths Matched to ITT Cannon","Career Summary","Detailed Professional Experience","Education, Certifications, Language","Compensation / Salary Information","Self-Introduction and Application Motivation")
    $found = @()
    foreach ($p in $doc.Paragraphs) {
        $t = CleanText $p
        foreach ($n in $names) {
            if ($t.StartsWith($n) -and ($t.Length -lt ($n.Length + 6))) { $found += [pscustomobject]@{ Pos=$p.Range.Start; Name=$n } }
        }
    }
    $ordered = $found | Sort-Object Pos
    $L += "ORDER:"
    $idx=0
    foreach ($f in $ordered) { $idx++; $L += ("  {0}. {1}" -f $idx, $f.Name) }

    $L += ("RESULT pages=" + $doc.ComputeStatistics(2) + " words=" + $doc.ComputeStatistics(0) + " tables=" + $doc.Tables.Count)
    $doc.Save(); $doc.ExportAsFixedFormat($pdf, 17); $doc.Close([ref]$true)
} finally {
    if ($word.Documents.Count -eq 0) { $word.Quit() }
}
$L | Out-File -FilePath $log -Encoding utf8
$L | ForEach-Object { Write-Output $_ }
Write-Output "MOVE-EDU DONE"
