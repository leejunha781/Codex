$ErrorActionPreference = "Stop"

function Normalize-Text {
    param([string]$Text)
    if ($null -eq $Text) { return "" }
    return (($Text -replace "[`r`n`a\x07]", " ") -replace "\s+", " ").Trim()
}

function Set-RangeTextKeepMark {
    param($Range, [string]$Text)
    $r = $Range.Duplicate
    if ($r.End -gt $r.Start) { $r.End = $r.End - 1 }
    $r.Text = $Text
}

function Find-ParagraphExact {
    param($Doc, [string]$Text)
    $needle = Normalize-Text $Text
    foreach ($p in $Doc.Paragraphs) {
        if ((Normalize-Text $p.Range.Text) -eq $needle) { return $p }
    }
    return $null
}

function Replace-ParagraphContains {
    param($Doc, [string]$Fragment, [string]$NewText)
    $needle = Normalize-Text $Fragment
    foreach ($p in $Doc.Paragraphs) {
        $txt = Normalize-Text $p.Range.Text
        if ($txt.Length -gt 0 -and $txt.Contains($needle)) {
            Set-RangeTextKeepMark -Range $p.Range -Text $NewText
            return $true
        }
    }
    return $false
}

function Compact-SelfIntro {
    param($Doc)
    $heading = Find-ParagraphExact -Doc $Doc -Text "Self-Introduction"
    if ($null -eq $heading) { throw "Self-Introduction heading not found" }

    Replace-ParagraphContains -Doc $Doc -Fragment "My goal is to help AVEVA grow meaningful Marine Engineering opportunities" -NewText "My goal is to help AVEVA grow meaningful Marine Engineering opportunities in Korea and Japan by combining marine-domain credibility, engineering-system understanding, competitive PLM insight and evidence-based customer engagement, while making AVEVA's value clear to customers who need practical answers to real shipbuilding lifecycle problems." | Out-Null

    foreach ($p in $Doc.Paragraphs) {
        if ($p.Range.Start -gt $heading.Range.Start) {
            $txt = Normalize-Text $p.Range.Text
            if ($txt.Length -gt 0) {
                $p.Range.Font.Name = "Arial"
                $p.Range.Font.Size = [single]10.5
                $p.Range.Font.Bold = [int]0
                $p.Range.ParagraphFormat.SpaceBefore = [single]0
                $p.Range.ParagraphFormat.SpaceAfter = [single]3
                $p.Range.ParagraphFormat.LineSpacingRule = 0
                if ($txt -match "^[1-4]\. ") {
                    $p.Range.Font.Bold = [int]1
                    $p.Range.Font.Size = [single]10.8
                    $p.Range.ParagraphFormat.SpaceBefore = [single]7
                    $p.Range.ParagraphFormat.SpaceAfter = [single]4
                }
            }
        }
    }
}

$mainDocx = "D:\이력서\AVEVA - Business Development Manager\Joonha_Lee_AVEVA_BDM_Engineering_Marine_Korea_Japan_Seoul_Resume.docx"
$mainPdf = "D:\이력서\AVEVA - Business Development Manager\Joonha_Lee_AVEVA_BDM_Engineering_Marine_Korea_Japan_Seoul_Resume.pdf"
$docx = "D:\이력서\AVEVA - Business Development Manager\Joonha_Lee_AVEVA_BDM_Engineering_Marine_Korea_Japan_Seoul_Resume_Formal_Tone.docx"
$pdf = "D:\이력서\AVEVA - Business Development Manager\Joonha_Lee_AVEVA_BDM_Engineering_Marine_Korea_Japan_Seoul_Resume_Formal_Tone.pdf"
$tempPdf = "C:\Users\namma\.claude\aveva_bdm_resume_work\self_intro_fit_export.pdf"

if (Test-Path -LiteralPath $tempPdf) { Remove-Item -LiteralPath $tempPdf -Force }

$word = New-Object -ComObject Word.Application
$word.Visible = $false
$doc = $null
try {
    $doc = $word.Documents.Open($docx, $false, $false)
    Compact-SelfIntro -Doc $doc
    $doc.Fields.Update() | Out-Null
    $doc.Save()
    $doc.ExportAsFixedFormat($tempPdf, 17)
    $pages = $doc.ComputeStatistics(2)
    $words = $doc.ComputeStatistics(0)
    $doc.Close([ref]$true)
    $doc = $null

    Copy-Item -LiteralPath $tempPdf -Destination $pdf -Force
    Copy-Item -LiteralPath $docx -Destination $mainDocx -Force
    Copy-Item -LiteralPath $pdf -Destination $mainPdf -Force

    "UPDATED_DOCX: $docx"
    "UPDATED_PDF: $pdf"
    "MAIN_DOCX: UPDATED"
    "MAIN_PDF: UPDATED"
    "PAGES: $pages"
    "WORDS: $words"
} finally {
    if ($doc -ne $null) {
        try { $doc.Close([ref]$false) } catch {}
    }
    if ($word -ne $null -and $word.Documents.Count -eq 0) {
        try { $word.Quit() } catch {}
    }
}
