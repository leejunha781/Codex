$ErrorActionPreference = "Stop"
$path = "D:\이력서\ITT Cannon\Joonha_Lee_ITT_Cannon_FAE_Resume_Final_Integrated_v2.docx"
$pdf  = "C:\Users\namma\.claude\itt_work\v2_current.pdf"
$log  = "C:\Users\namma\.claude\itt_work\move_interest_log.txt"
$L = @()
function CleanText($p) { return ($p.Range.Text -replace "[\r\n\a\x07]","" -replace "\s+"," ").Trim() }

$word = New-Object -ComObject Word.Application
$word.Visible = $false
try {
    $doc = $word.Documents.Open($path, $false, $false)

    # 1) Rename Profile heading -> "Professional Profile"
    foreach ($p in $doc.Paragraphs) {
        if ((CleanText $p).StartsWith("Professional Profile and Interest")) {
            $r = $p.Range.Duplicate; $r.End = $r.End - 1; $r.Text = "Professional Profile"
            $L += "1) Heading renamed to 'Professional Profile'."
            break
        }
    }

    # 2) Delete the redundant Profile 'interest' paragraph (unique phrase: 'built my career around')
    $deleted = $false
    foreach ($p in $doc.Paragraphs) {
        $t = CleanText $p
        if ($t.Contains("I am interested in ITT Cannon Connector Korea") -and $t.Contains("built my career around")) {
            $p.Range.Delete()
            $deleted = $true
            $L += "2) Removed redundant Profile interest paragraph."
            break
        }
    }
    $L += "   deleted=$deleted"

    # ---- verify ----
    $cntInterest = 0; $cntCareerAround = 0; $hasNewHeading = $false; $hasOldHeading = $false
    $expect = @("Professional Profile","Key Strengths Matched to ITT Cannon","Education, Certifications, Language","Career Summary","Detailed Professional Experience","Compensation / Salary Information","Self-Introduction and Application Motivation","Reason for Interest in ITT Cannon FAE Role")
    $missing = @()
    foreach ($e in $expect) {
        $ok=$false; foreach ($p in $doc.Paragraphs) { if ((CleanText $p).StartsWith($e)) { $ok=$true; break } }
        if (-not $ok) { $missing += $e }
    }
    foreach ($p in $doc.Paragraphs) {
        $t = CleanText $p
        if ($t.Contains("I am interested in ITT Cannon Connector Korea")) { $cntInterest++ }
        if ($t.Contains("built my career around")) { $cntCareerAround++ }
        if ($t.StartsWith("Professional Profile and Interest")) { $hasOldHeading = $true }
        if ($t -eq "Professional Profile") { $hasNewHeading = $true }
    }
    $L += "VERIFY interestParas=$cntInterest (expect 1, in Self-Intro) ; careerAroundParas=$cntCareerAround (expect 0)"
    $L += "VERIFY newHeading=$hasNewHeading oldHeading=$hasOldHeading"
    $L += ("VERIFY missingHeadings=" + ($(if ($missing.Count -eq 0) { 'none' } else { $missing -join ', ' })))
    $L += ("RESULT pages=" + $doc.ComputeStatistics(2) + " words=" + $doc.ComputeStatistics(0))

    $doc.Save(); $doc.ExportAsFixedFormat($pdf, 17); $doc.Close([ref]$true)
} finally {
    if ($word.Documents.Count -eq 0) { $word.Quit() }
}
$L | Out-File -FilePath $log -Encoding utf8
$L | ForEach-Object { Write-Output $_ }
Write-Output "MOVE-INTEREST DONE"
