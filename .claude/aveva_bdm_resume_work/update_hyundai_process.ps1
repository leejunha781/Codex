$ErrorActionPreference = "Stop"

function Set-RangeTextKeepMark {
    param(
        $Range,
        [string]$Text
    )
    $r = $Range.Duplicate
    if ($r.End -gt $r.Start) {
        $r.End = $r.End - 1
    }
    $r.Text = $Text
}

function Bold-Substring {
    param(
        $Range,
        [string]$Needle
    )
    $txt = $Range.Text
    $idx = $txt.IndexOf($Needle)
    if ($idx -ge 0) {
        $sub = $Range.Duplicate
        $sub.Start = $Range.Start + $idx
        $sub.End = $sub.Start + $Needle.Length
        $sub.Font.Bold = [int]1
    }
}

$docx = "D:\이력서\AVEVA - Business Development Manager\Joonha_Lee_AVEVA_BDM_Engineering_Marine_Korea_Japan_Seoul_Resume.docx"
$pdf = "D:\이력서\AVEVA - Business Development Manager\Joonha_Lee_AVEVA_BDM_Engineering_Marine_Korea_Japan_Seoul_Resume.pdf"

$word = New-Object -ComObject Word.Application
$word.Visible = $false
$doc = $null
try {
    $doc = $word.Documents.Open($docx, $false, $false)

    $table = $doc.Tables.Item(3)
    $cell = $table.Range.Cells.Item(6)
    $newText = "ROKN ship/submarine programmes plus working understanding of HD Hyundai-style standard shipbuilding process: requirements -> basic/detail design -> E3D/Marine/Draw -> MTO/eBOM -> hull/block effectivity & baseline -> ECR/ECO -> production BOM/release -> commissioning/handover evidence."

    Set-RangeTextKeepMark -Range $cell.Range -Text $newText
    $cell.Range.Font.Name = "Arial"
    $cell.Range.Font.Size = [single]10
    $cell.Range.Font.Bold = [int]0
    $cell.Range.ParagraphFormat.SpaceBefore = [single]0
    $cell.Range.ParagraphFormat.SpaceAfter = [single]0
    $cell.Range.ParagraphFormat.LineSpacingRule = 0
    Bold-Substring -Range $cell.Range -Needle "HD Hyundai-style standard shipbuilding process"

    $doc.Fields.Update() | Out-Null
    $doc.Save()
    $doc.ExportAsFixedFormat($pdf, 17)
    $pages = $doc.ComputeStatistics(2)
    $words = $doc.ComputeStatistics(0)
    $doc.Close([ref]$true)
    $doc = $null
    "UPDATED: $docx"
    "PDF: $pdf"
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
