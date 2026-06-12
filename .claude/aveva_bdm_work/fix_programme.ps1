# Replace British "programme(s)" with business-standard "program(s)" in both resume files
# Word-level inline replacement (preserves surrounding bold/format); Find.Execute is unusable from PS.
$ErrorActionPreference = 'Stop'

$files = @(
    'D:\이력서\AVEVA - Business Development Manager\Joonha_Lee_AVEVA_BDM_Engineering_Marine_Korea_Japan_Seoul_Resume_Formal_Tone.docx',
    'D:\이력서\AVEVA - Business Development Manager\Joonha_Lee_AVEVA_BDM_Engineering_Marine_Korea_Japan_Seoul_Resume.docx'
)

$word = New-Object -ComObject Word.Application
$word.Visible = $false
try {
    foreach ($f in $files) {
        Write-Output "=== $([System.IO.Path]::GetFileName($f))"
        $doc = $word.Documents.Open($f, $false, $false)
        try {
            $replaced = 0; $fails = 0
            foreach ($p in $doc.Paragraphs) {
                if ($p.Range.Text -notmatch '(?i)programme') { continue }
                $guard = 0
                while ($guard -lt 25) {
                    $t = $p.Range.Text
                    $m = [regex]::Match($t, '(?i)programme')
                    if (-not $m.Success) { break }
                    $rep = 'program'
                    if ($m.Value[0] -ceq 'P') { $rep = 'Program' }
                    $start = $p.Range.Start + $m.Index
                    $done = $false
                    for ($d = 0; $d -le 3 -and -not $done; $d = $(if ($d -le 0) { -$d + 1 } else { -$d })) {
                        # try offsets 0, +1, -1, +2, -2, +3, -3
                        $r = $doc.Range($start + $d, $start + $d + 9)
                        if ($r.Text -ceq $m.Value) {
                            $r.Text = $rep
                            $done = $true
                        }
                        if ($d -eq -3) { break }
                    }
                    if ($done) { $replaced++ } else { $fails++; Write-Output ("  OFFSET FAIL near: " + ($t.Substring([Math]::Max(0,$m.Index-20), [Math]::Min(50, $t.Length - [Math]::Max(0,$m.Index-20))) -replace "[\r\a]",'')); break }
                    $guard++
                }
            }
            # final verification
            $left = 0
            foreach ($p in $doc.Paragraphs) {
                $left += ([regex]::Matches($p.Range.Text, '(?i)programme')).Count
            }
            Write-Output "  replaced: $replaced | failures: $fails | 'programme' remaining: $left"
            $doc.Save()
            $pdf = [System.IO.Path]::ChangeExtension($f, 'pdf')
            $doc.ExportAsFixedFormat($pdf, 17)
            Write-Output "  Saved + PDF re-exported (pages: $($doc.ComputeStatistics(2)))"
        } finally {
            $doc.Close([ref]$false)
        }
    }
} finally {
    if ($word.Documents.Count -eq 0) { $word.Quit() }
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($word)
}
Write-Output 'PROGRAMME FIX DONE'
