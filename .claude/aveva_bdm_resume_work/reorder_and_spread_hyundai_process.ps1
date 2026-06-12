$ErrorActionPreference = "Stop"

function Normalize-Text {
    param([string]$Text)
    if ($null -eq $Text) { return "" }
    return (($Text -replace "[`r`n`a\x07]", " ") -replace "\s+", " ").Trim()
}

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

function Set-CellText {
    param(
        $Table,
        [int]$Index,
        [string]$Text,
        [bool]$Bold = $false,
        [single]$Size = 10.5
    )
    $cell = $Table.Range.Cells.Item($Index)
    Set-RangeTextKeepMark -Range $cell.Range -Text $Text
    $cell.Range.Font.Name = "Arial"
    $cell.Range.Font.Size = [single]$Size
    $cell.Range.Font.Bold = [int]0
    if ($Bold) { $cell.Range.Font.Bold = [int]1 }
    $cell.Range.ParagraphFormat.SpaceBefore = [single]0
    $cell.Range.ParagraphFormat.SpaceAfter = [single]0
    $cell.Range.ParagraphFormat.LineSpacingRule = 0
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

function Replace-ParagraphContains {
    param(
        $Doc,
        [string]$Fragment,
        [string]$NewText,
        [bool]$Bold = $false
    )
    $needle = Normalize-Text $Fragment
    $hits = 0
    foreach ($p in $Doc.Paragraphs) {
        $txt = Normalize-Text $p.Range.Text
        if ($txt.Length -gt 0 -and $txt.Contains($needle)) {
            Set-RangeTextKeepMark -Range $p.Range -Text $NewText
            $p.Range.Font.Name = "Arial"
            $p.Range.Font.Size = [single]11
            $p.Range.Font.Bold = [int]0
            if ($Bold) { $p.Range.Font.Bold = [int]1 }
            $p.Range.ParagraphFormat.SpaceBefore = [single]0
            $p.Range.ParagraphFormat.SpaceAfter = [single]8
            $hits++
        }
    }
    return $hits
}

function Find-ParagraphExact {
    param(
        $Doc,
        [string]$Text
    )
    $needle = Normalize-Text $Text
    foreach ($p in $Doc.Paragraphs) {
        if ((Normalize-Text $p.Range.Text) -eq $needle) {
            return $p
        }
    }
    return $null
}

$docx = "D:\이력서\AVEVA - Business Development Manager\Joonha_Lee_AVEVA_BDM_Engineering_Marine_Korea_Japan_Seoul_Resume.docx"
$pdf = "D:\이력서\AVEVA - Business Development Manager\Joonha_Lee_AVEVA_BDM_Engineering_Marine_Korea_Japan_Seoul_Resume.pdf"

$word = New-Object -ComObject Word.Application
$word.Visible = $false
$doc = $null
try {
    $doc = $word.Documents.Open($docx, $false, $false)

    # Best front-page flow for this BDM application:
    # Core Competencies -> Career Summary -> Job-Fit Map -> Education -> Detailed Experience.
    $education = Find-ParagraphExact -Doc $doc -Text "Education"
    $career = Find-ParagraphExact -Doc $doc -Text "Career Summary"
    if ($education -ne $null -and $career -ne $null -and $education.Range.Start -lt $career.Range.Start) {
        $eduBlock = $doc.Range($education.Range.Start, $career.Range.Start)
        $eduBlock.Cut() | Out-Null
        $target = Find-ParagraphExact -Doc $doc -Text "Detailed Professional Experience"
        if ($target -ne $null) {
            $pasteRange = $doc.Range($target.Range.Start, $target.Range.Start)
            $pasteRange.Collapse(1)
            $pasteRange.Paste()
        }
    }

    # Spread HD Hyundai standard-process understanding in the highest-value places.
    $coreCell = $doc.Tables.Item(3).Range.Cells.Item(6)
    Bold-Substring -Range $coreCell.Range -Needle "HD Hyundai standard process"

    Set-CellText -Table $doc.Tables.Item(4) -Index 20 -Text "15 years Korea naval/shipbuilding customer delivery; shipyard pain-point discovery plus HD Hyundai standard process awareness from requirements/VCRM to design, MTO/eBOM, ECR/ECO and handover evidence." -Size ([single]10.5)

    Set-CellText -Table $doc.Tables.Item(5) -Index 8 -Text "Worked 15 years in ROKN/shipyard programmes; mapped HD Hyundai standard process from requirements/VCRM through E3D/Marine/Draw, MTO/eBOM, effectivity/baseline, ECR/ECO and handover evidence; benchmarked AVEVA against Siemens, Dassault, Hexagon, PTC and NAPA." -Size ([single]10)
    Set-CellText -Table $doc.Tables.Item(5) -Index 9 -Text "Gives the sales team process-grounded credibility and a differentiated AVEVA narrative for shipbuilding, offshore and naval customer conversations." -Size ([single]10)

    Replace-ParagraphContains -Doc $doc -Fragment "Delivered ROK Navy ship and submarine communication systems, shipboard broadcasting systems, internal networks, CCTV/VOD and C4I-linked systems" -NewText "Delivered ROK Navy ship and submarine communication systems, shipboard broadcasting systems, internal networks, CCTV/VOD and C4I-linked systems from requirements analysis through design, installation, commissioning, acceptance and after-delivery support. I also mapped this experience to the HD Hyundai standard shipbuilding flow: requirements/VCRM, basic/detail design, E3D/Marine/Draw, MTO/eBOM, hull/block effectivity and baseline, ECR/ECO, production release and handover evidence. This is the strongest foundation for AVEVA Marine BDM because it gives practical understanding of Korean shipyard workflows, naval customer language, system integration pain points and lifecycle documentation needs." | Out-Null

    Replace-ParagraphContains -Doc $doc -Fragment "At Daeyang Electric, I managed and supported naval communication and shipboard systems from requirements analysis" -NewText "At Daeyang Electric, I managed and supported naval communication and shipboard systems from requirements analysis through proposal, design, installation, FAT, SAT, sea trial, customer acceptance and after-delivery support. I converted Navy and shipyard requirements into system architecture, equipment configuration, ICDs, test plans, test procedures, acceptance criteria and close-out documents. I can now connect that experience to the HD Hyundai standard process from requirements/VCRM through E3D/Marine/Draw, MTO/eBOM, effectivity/baseline, ECR/ECO and handover evidence, which gives me credible working knowledge for Korean marine customer engagement." | Out-Null

    Replace-ParagraphContains -Doc $doc -Fragment "The future strategy I would explain to marine customers is an open AVEVA decision control plane" -NewText "The future strategy I would explain to marine customers is an open AVEVA decision control plane fitted to the large-shipyard standard process: configuration, impact, evidence, approval and operations feedback connected through customer-owned APIs, governed workflows, policy gates and trusted evidence. This approach keeps specialist tools where they are strong, while allowing AVEVA to own the cross-domain decision loop for design change, assurance evidence, production readiness, handover and lifecycle learning." | Out-Null

    # Tighten table fonts after added process language.
    foreach ($idx in @(4,5)) {
        foreach ($cell in $doc.Tables.Item($idx).Range.Cells) {
            $cell.Range.ParagraphFormat.SpaceBefore = [single]0
            $cell.Range.ParagraphFormat.SpaceAfter = [single]0
            $cell.Range.ParagraphFormat.LineSpacingRule = 0
        }
    }

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
