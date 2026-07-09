#requires -Version 5.1
<#
.SYNOPSIS
  Convert freelance markdown portfolio samples to DOCX and PDF via Word COM.
#>
param(
    [string]$SamplesDir = "C:\Users\namma\freelance\02-samples",
    [string]$OutputDir = "C:\Users\namma\freelance\02-samples\pdf"
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = "Stop"

if (-not (Test-Path -LiteralPath $OutputDir)) {
    New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null
}

function Get-WordStyleName {
    param([int]$Level)
    $candidates = @(
        "Heading $Level",
        "제목 $Level",
        "标题 $Level"
    )
    return $candidates[0]
}

function Set-ParagraphHeading {
    param($Paragraph, [int]$Level)
    $sizes = @{ 1 = 16; 2 = 13; 3 = 12 }
    $Paragraph.Range.Font.Bold = $true
    $Paragraph.Range.Font.Size = $sizes[$Level]
    $Paragraph.Range.Font.Name = "Calibri"
    $Paragraph.SpaceAfter = 6
}

function Add-TextLine {
    param(
        [Parameter(Mandatory = $true)]$Document,
        [Parameter(Mandatory = $true)][string]$Text,
        [int]$HeadingLevel = 0
    )

    if ([string]::IsNullOrWhiteSpace($Text)) {
        return
    }

    $range = $Document.Content
    $range.Collapse(0)
    $range.Text = ($Text -replace '\*\*(.+?)\*\*', '$1') + "`r"
    $para = $Document.Paragraphs.Last

    if ($HeadingLevel -gt 0) {
        Set-ParagraphHeading -Paragraph $para -Level $HeadingLevel
    }
    else {
        $para.Range.Font.Name = "Calibri"
        $para.Range.Font.Size = 11
        $para.Range.Font.Bold = $false
    }
}

function Add-MarkdownTableToWord {
    param(
        [Parameter(Mandatory = $true)]$Document,
        [Parameter(Mandatory = $true)][string[]]$TableLines
    )

    $rows = New-Object System.Collections.Generic.List[object]
    foreach ($line in $TableLines) {
        if ($line -match '^\|') {
            $cells = ($line.Trim('|') -split '\|') | ForEach-Object { $_.Trim() }
            if ($cells -join '' -match '^[-:\s]+$') { continue }
            $rows.Add($cells)
        }
    }

    if ($rows.Count -eq 0) { return }

    $colCount = ($rows | ForEach-Object { $_.Count } | Measure-Object -Maximum).Maximum
    $range = $Document.Content
    $range.Collapse(0)
    $table = $Document.Tables.Add($range, $rows.Count, $colCount)
    try {
        $table.Style = "표 간단1"
    }
    catch {
        try { $table.Style = "Table Grid" } catch {}
    }
    $table.Borders.Enable = $true

    for ($r = 0; $r -lt $rows.Count; $r++) {
        for ($c = 0; $c -lt $rows[$r].Count; $c++) {
            $cell = $table.Cell($r + 1, $c + 1)
            $cell.Range.Text = $rows[$r][$c]
            if ($r -eq 0) {
                $cell.Range.Font.Bold = $true
            }
        }
    }

    $Document.Content.InsertAfter("`r")
}

function Convert-MarkdownToDocx {
    param(
        [Parameter(Mandatory = $true)][string]$MarkdownPath,
        [Parameter(Mandatory = $true)][string]$DocxPath
    )

    $lines = Get-Content -LiteralPath $MarkdownPath -Encoding UTF8
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $doc = $null

    try {
        $doc = $word.Documents.Add()
        $doc.Content.Text = ""

        $i = 0
        while ($i -lt $lines.Count) {
            $line = $lines[$i]

            if ($line -match '^\|') {
                $tableLines = @()
                while ($i -lt $lines.Count -and $lines[$i] -match '^\|') {
                    $tableLines += $lines[$i]
                    $i++
                }
                Add-MarkdownTableToWord -Document $doc -TableLines $tableLines
                continue
            }

            if ($line -eq '---') {
                $i++
                continue
            }

            if ($line -match '^\#{1}\s+(.+)$') {
                Add-TextLine -Document $doc -Text $Matches[1] -HeadingLevel 1
            }
            elseif ($line -match '^\#{2}\s+(.+)$') {
                Add-TextLine -Document $doc -Text $Matches[1] -HeadingLevel 2
            }
            elseif ($line -match '^\#{3}\s+(.+)$') {
                Add-TextLine -Document $doc -Text $Matches[1] -HeadingLevel 3
            }
            elseif (-not [string]::IsNullOrWhiteSpace($line)) {
                Add-TextLine -Document $doc -Text $line
            }

            $i++
        }

        if (Test-Path -LiteralPath $DocxPath) {
            Remove-Item -LiteralPath $DocxPath -Force
        }

        $doc.SaveAs2([ref]$DocxPath, [ref]16)
        $doc.Close([ref]$false)
        $doc = $null
    }
    finally {
        if ($null -ne $doc) {
            try { $doc.Close([ref]$false) } catch {}
        }
        if ($null -ne $word -and $word.Documents.Count -eq 0) {
            $word.Quit()
        }
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($word)
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

function Export-DocxToPdf {
    param(
        [Parameter(Mandatory = $true)][string]$DocxPath,
        [Parameter(Mandatory = $true)][string]$PdfPath
    )

    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $doc = $null
    try {
        $doc = $word.Documents.Open($DocxPath, $false, $true)
        if (Test-Path -LiteralPath $PdfPath) {
            Remove-Item -LiteralPath $PdfPath -Force
        }
        $doc.ExportAsFixedFormat($PdfPath, 17)
        $doc.Close([ref]$false)
        $doc = $null
    }
    finally {
        if ($null -ne $doc) {
            try { $doc.Close([ref]$false) } catch {}
        }
        if ($null -ne $word -and $word.Documents.Count -eq 0) {
            $word.Quit()
        }
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($word)
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

$samples = @(
    "RCA_Report_Sample_Satellite_Terminal_Link_Loss.md",
    "FAT_SAT_Checklist_Sample_LEO_Terminal.md",
    "PLM_Concept_Deck_Outline_Sample.md"
)

$results = @()
foreach ($name in $samples) {
    $mdPath = Join-Path $SamplesDir $name
    if (-not (Test-Path -LiteralPath $mdPath)) {
        Write-Warning "Missing: $mdPath"
        continue
    }

    $base = [System.IO.Path]::GetFileNameWithoutExtension($name)
    $docxPath = Join-Path $OutputDir ($base + ".docx")
    $pdfPath = Join-Path $OutputDir ($base + ".pdf")

    Write-Host "Converting $name ..."
    Convert-MarkdownToDocx -MarkdownPath $mdPath -DocxPath $docxPath
    Export-DocxToPdf -DocxPath $docxPath -PdfPath $pdfPath

    $results += [PSCustomObject]@{
        Source  = $name
        Docx    = $docxPath
        Pdf     = $pdfPath
        PdfSize = (Get-Item -LiteralPath $pdfPath).Length
    }
}

Write-Host ""
Write-Host "=== Conversion complete ==="
$results | Format-Table -AutoSize
