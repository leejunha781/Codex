$ErrorActionPreference = "Stop"

$DocPath = "D:\이력서\AVEVA - Business Development Manager\Joonha_Lee_ABB_Portfolio_Industry_Manager_Marine_Ports_Industrial_Automation_Resume.docx"
$PdfPath = "D:\이력서\AVEVA - Business Development Manager\Joonha_Lee_ABB_Portfolio_Industry_Manager_Marine_Ports_Industrial_Automation_Resume.pdf"
$OldHeader = "Joonha Lee | AVEVA BDM Engineering (Marine) | Resume"
$NewHeader = "Joonha Lee | ABB Portfolio & Industry Manager | Resume"

function Normalize-Text([string]$Text) {
    return (($Text -replace "[\r\n\a\x07]", "") -replace "\s+", " ").Trim()
}

function Set-RangeTextKeepMark($Range, [string]$Text) {
    $r = $Range.Duplicate
    if ($r.End -gt $r.Start) {
        $r.End = $r.End - 1
    }
    $r.Text = $Text
}

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
$doc = $null

try {
    $doc = $word.Documents.Open($DocPath, $false, $false)

    $headerHits = 0
    foreach ($section in $doc.Sections) {
        foreach ($header in $section.Headers) {
            foreach ($p in $header.Range.Paragraphs) {
                if ((Normalize-Text $p.Range.Text) -eq $OldHeader) {
                    Set-RangeTextKeepMark $p.Range $NewHeader
                    $headerHits++
                }
            }
            foreach ($shape in $header.Shapes) {
                if ($shape.TextFrame.HasText -ne 0) {
                    if ((Normalize-Text $shape.TextFrame.TextRange.Text) -eq $OldHeader) {
                        $shape.TextFrame.TextRange.Text = $NewHeader
                        $headerHits++
                    }
                }
            }
        }
    }

    $doc.Save()
    $doc.ExportAsFixedFormat($PdfPath, 17)

    $allText = New-Object System.Text.StringBuilder
    $story = $doc.StoryRanges
    foreach ($range in $story) {
        [void]$allText.AppendLine($range.Text)
    }
    foreach ($section in $doc.Sections) {
        foreach ($header in $section.Headers) {
            [void]$allText.AppendLine($header.Range.Text)
        }
        foreach ($footer in $section.Footers) {
            [void]$allText.AppendLine($footer.Range.Text)
        }
    }

    $text = $allText.ToString()
    $terms = @("AVEVA", "ABB", "Portfolio Positioning", "Industry Strategy", "Solution Portfolio Mapping", "Technical Sales Enablement", "Mission-Critical", "Industrial Automation", "Power & Energy Infrastructure")
    foreach ($term in $terms) {
        $count = ([regex]::Matches($text, [regex]::Escape($term))).Count
        Write-Host ("TERM {0}: {1}" -f $term, $count)
    }

    Write-Host "HeaderHits: $headerHits"
    Write-Host "Pages     : $($doc.ComputeStatistics(2))"
    Write-Host "Words     : $($doc.ComputeStatistics(0))"
    $doc.Close([ref]$true)
    $doc = $null
}
finally {
    if ($doc -ne $null) {
        try { $doc.Close([ref]$false) } catch {}
    }
    if ($word -ne $null -and $word.Documents.Count -eq 0) {
        $word.Quit()
    }
}
