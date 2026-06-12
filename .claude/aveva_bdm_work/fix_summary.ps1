# Polish the Professional Summary paragraph in all three resume files
$ErrorActionPreference = 'Stop'

$prefix = 'Marine and industrial engineering professional with'
$new = 'Marine and industrial engineering professional with over 21 years of experience across ROK Navy shipbuilding programmes (FFX Batch-II, Jangbogo-III), defence electronics and a global LEO satellite terminal programme. Combines hands-on shipyard lifecycle credibility — requirements/VCRM, E3D/Marine design data, MTO/eBOM, baseline/effectivity control, ECR/ECO and class/handover evidence — with technical value selling that translates marine customer pain points, such as design rework and system integration challenges, into value propositions, benchmark-based benefit cases and qualified opportunities. Ready to support AVEVA regional sales teams, Account Managers and Pre-Sales Consultants across Korea and Japan with marine domain credibility, competitive PLM insight covering Siemens, Dassault, Hexagon, PTC and NAPA, and structured pipeline support aligned with Salesforce CRM.'

function VisText([string]$t) { ($t -replace "[\r\n\a\x0b\x07]", '').TrimEnd() }

$files = @(
    'D:\이력서\AVEVA - Business Development Manager\Joonha_Lee_AVEVA_BDM_Engineering_Marine_Korea_Japan_Seoul_Resume_Formal_Tone.docx',
    'D:\이력서\AVEVA - Business Development Manager\Joonha_Lee_AVEVA_BDM_Engineering_Marine_Korea_Japan_Seoul_Resume.docx',
    'D:\이력서\AVEVA - Business Development Manager\Joonha_Lee_AVEVA_BDM_Engineering_Marine_Korea_Japan_Seoul_Resume_ATS_2Page.docx'
)

$word = New-Object -ComObject Word.Application
$word.Visible = $false
try {
    foreach ($f in $files) {
        Write-Output "=== $([System.IO.Path]::GetFileName($f))"
        $doc = $word.Documents.Open($f, $false, $false)
        try {
            $done = $false
            foreach ($p in $doc.Paragraphs) {
                $vis = VisText $p.Range.Text
                if (-not $vis.StartsWith($prefix)) { continue }
                # replace with End-1, then repair any stray tail chars position-safely
                $r = $p.Range.Duplicate
                $r.End = $r.End - 1
                $r.Text = $new
                $guard = 0
                while ($guard -lt 8) {
                    $cur = VisText $p.Range.Text
                    if ($cur -eq $new) { break }
                    if (-not $cur.StartsWith($new)) { Write-Output "  UNEXPECTED: ...$($cur.Substring([Math]::Max(0,$cur.Length-50)))"; break }
                    $chars = $p.Range.Characters
                    $i = $chars.Count
                    while ($i -ge 1) {
                        $code = [int][char]($chars.Item($i).Text[0])
                        if ($code -ne 13 -and $code -ne 10 -and $code -ne 7 -and $code -ne 11) { break }
                        $i--
                    }
                    if ($i -lt 1) { break }
                    $chars.Item($i).Delete() | Out-Null
                    $guard++
                }
                $final = VisText $p.Range.Text
                if ($final -eq $new) { Write-Output "  Summary replaced + verified exact (repaired $guard stray char(s))."; $done = $true }
                else { Write-Output "  VERIFY FAILED: ...$($final.Substring([Math]::Max(0,$final.Length-60)))" }
                break
            }
            if (-not $done) { Write-Output '  WARNING: summary paragraph not found/verified!' }
            $doc.Save()
            $pdf = [System.IO.Path]::ChangeExtension($f, 'pdf')
            $doc.ExportAsFixedFormat($pdf, 17)
            Write-Output "  Pages: $($doc.ComputeStatistics(2)) | PDF re-exported."
        } finally {
            $doc.Close([ref]$false)
        }
    }
} finally {
    if ($word.Documents.Count -eq 0) { $word.Quit() }
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($word)
}
Write-Output 'SUMMARY FIX DONE'
