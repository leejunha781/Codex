$ErrorActionPreference = "Stop"

$pdf = "D:\이력서\AVEVA - Business Development Manager\Joonha_Lee_AVEVA_BDM_Engineering_Marine_Korea_Japan_Seoul_Resume.pdf"
$out = "C:\Users\namma\.claude\aveva_bdm_resume_work\bdm_pdf_text_check.txt"

$word = New-Object -ComObject Word.Application
$word.Visible = $false
$doc = $null
try {
    $doc = $word.Documents.Open($pdf, $false, $true)
    $text = $doc.Content.Text
    $checks = @(
        "Business Development Manager - Engineering (Marine), Korea / Japan",
        "Application Location: Seoul, Korea",
        "Marine Business Development / Pipeline Support",
        "Salesforce CRM pipeline discipline",
        "Impact, Aspiration, Curiosity, and Trust",
        "open AVEVA decision control plane",
        "Siemens Teamcenter",
        "Dassault 3DEXPERIENCE",
        "Hexagon",
        "PTC",
        "NAPA"
    )
    $lines = New-Object System.Collections.Generic.List[string]
    $lines.Add("PDF: $pdf")
    $lines.Add("TEXT_LENGTH: " + $text.Length)
    foreach ($check in $checks) {
        $present = $text.Contains($check)
        $lines.Add(("CHECK {0}: {1}" -f $present, $check))
    }
    foreach ($old in @("Marine Principal Technical Support & Consultant", "PLM SME, Busan")) {
        $present = $text.Contains($old)
        $lines.Add(("OLD_TERM_PRESENT {0}: {1}" -f $present, $old))
    }
    [System.IO.File]::WriteAllLines($out, $lines, [System.Text.Encoding]::UTF8)
    $doc.Close([ref]$false)
    $doc = $null
    $lines
} finally {
    if ($doc -ne $null) {
        try { $doc.Close([ref]$false) } catch {}
    }
    if ($word -ne $null -and $word.Documents.Count -eq 0) {
        try { $word.Quit() } catch {}
    }
}
