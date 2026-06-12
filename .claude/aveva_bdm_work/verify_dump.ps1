$ErrorActionPreference = 'Stop'
$word = New-Object -ComObject Word.Application
$word.Visible = $false
try {
    $doc = $word.Documents.Open('D:\이력서\AVEVA - Business Development Manager\Joonha_Lee_AVEVA_BDM_Engineering_Marine_Korea_Japan_Seoul_Resume_Formal_Tone.docx', $false, $true)
    try {
        $sb = New-Object System.Text.StringBuilder
        $i = 0
        foreach ($p in $doc.Paragraphs) {
            $i++
            $txt = ($p.Range.Text -replace "[\r\n\a\x0b\x07]", '').TrimEnd()
            [void]$sb.AppendLine(("[{0:d3}] {1}" -f $i, $txt))
        }
        [System.IO.File]::WriteAllText('C:\Users\namma\.claude\aveva_bdm_work\resume_after.txt', $sb.ToString(), [System.Text.Encoding]::UTF8)
        Write-Output "Dumped $i paragraphs; pages: $($doc.ComputeStatistics(2))"
    } finally { $doc.Close($false) }
} finally {
    if ($word.Documents.Count -eq 0) { $word.Quit() }
}
$old = @('XecureHiteRon','Exicure','Nesslab','Senior Manager, SIT','IELTS','V18','consultant role','HD Hyundai-standard','Leejunha781','Thread preparation')
$new = @('Excure Hitron','NesLab','General Manager, SIT','Professional working English','Industrial AI Assistant','white space','Salesforce CRM-ready','Professional Summary','thesis research','HART/Modbus','Deputy General Manager, Naval System Engineer','class/regulatory evidence','tender conditions')
$txt = [System.IO.File]::ReadAllText('C:\Users\namma\.claude\aveva_bdm_work\resume_after.txt')
Write-Output '--- OLD strings (should be 0 hits):'
foreach ($s in $old) { $c = ([regex]::Matches($txt, [regex]::Escape($s))).Count; Write-Output "  [$c] $s" }
Write-Output '--- NEW strings (should be >=1):'
foreach ($s in $new) { $c = ([regex]::Matches($txt, [regex]::Escape($s))).Count; Write-Output "  [$c] $s" }
