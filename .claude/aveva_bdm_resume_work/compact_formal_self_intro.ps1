$ErrorActionPreference = "Stop"

function Normalize-Text {
    param([string]$Text)
    if ($null -eq $Text) { return "" }
    return (($Text -replace "[`r`n`a\x07]", " ") -replace "\s+", " ").Trim()
}

function Find-ParagraphExact {
    param($Doc, [string]$Text)
    $needle = Normalize-Text $Text
    foreach ($p in $Doc.Paragraphs) {
        if ((Normalize-Text $p.Range.Text) -eq $needle) { return $p }
    }
    return $null
}

function Replace-SelfIntroductionCompact {
    param($Doc)
    $heading = Find-ParagraphExact -Doc $Doc -Text "Self-Introduction"
    if ($null -eq $heading) { throw "Self-Introduction heading not found" }

    $paras = @(
        "1. Motivation for Application",
        "I am applying for AVEVA's Business Development Manager - Engineering (Marine), Korea / Japan role in Seoul because it sits at the intersection of marine engineering, customer value and industrial software. My career has been built in environments where customer requirements, technical risk, verification evidence and delivery commitments must be converted into practical decisions.",
        "For more than 21 years, I have worked across naval shipbuilding, defence electronics and satellite communication systems. The strongest foundation is my 15 years at Daeyang Electric, where I supported ROK Navy ship and submarine programmes including FFX Batch-II and Jangbogo/KSS-III, coordinating with Navy users, shipyards, inspectors, subcontractors, engineering, production, quality and project teams.",
        "Marine customers do not buy software functions alone. They need lower rework, stronger configuration control, faster change-impact review, reliable approval evidence and a trusted digital thread from design through build, handover and operations. That is the business value I see in AVEVA Engineering / Marine.",
        "AVEVA's values of Impact, Aspiration, Curiosity and Trust are close to my working style. I aim for practical impact, ask why before deciding, and build trust through evidence, clear follow-up and accountable communication. I also respect AVEVA's sustainability direction: better engineering data should reduce waste, improve productivity and help customers use industrial resources more responsibly.",
        "2. Fit for the Role",
        "My value is the ability to translate shipbuilding reality into business-development action. I am comfortable listening to engineers, identifying the real issue behind a technical symptom, and turning it into a discovery question, value proposition, benchmark assumption, PoC scope, risk item or qualified opportunity.",
        "At Daeyang Electric, I converted Navy and shipyard requirements into system architecture, equipment configuration, ICDs, test plans, procedures, acceptance criteria and close-out documents. That experience connects naturally with the HD Hyundai-standard process: requirements/VCRM, E3D/Marine/Draw, MTO/eBOM, baseline/effectivity, ECR/ECO and handover evidence.",
        "At Intellian Technologies, I worked in a global customer environment, managing requirements, field issues, defect history, ECO/BOM changes, verification evidence and release judgement for a LEO satellite-terminal programme. Managing more than 400 verification checklist items reinforced one lesson: customer trust is earned through evidence, not claims.",
        "Commercially, this background fits the technical value-selling side of the role: customer discovery, proposal support, benefit explanation, risk reduction, lifecycle data thinking and cross-functional follow-up. I would support Account Managers and Pre-Sales Consultants with CRM-aware opportunity qualification, discovery evidence, value-proposition inputs and disciplined next actions.",
        "3. AVEVA Marine PLM Perspective",
        "While preparing for the AVEVA Marine PLM consultant role, I analysed AVEVA Marine / PLM from the perspectives of product position, competitor strengths, customer pain points and future strategy. My conclusion from the V18 meeting deck was clear: AVEVA should not compete only as a closed PLM backbone. Its stronger position is an engineering-to-operations decision layer above specialist tools.",
        "My comparison view is practical. Siemens Teamcenter is strong in configuration, BOM and change governance; Dassault 3DEXPERIENCE in collaboration and virtual-twin positioning; Hexagon in asset lifecycle and operations; PTC in requirements, ALM and traceability; and NAPA in naval architecture and stability calculation. AVEVA's opportunity is to connect engineering truth, lifecycle context and operations feedback through E3D/Marine, Unified Engineering, AIM, PI/CONNECT and open integration.",
        "The customer message I would carry is an open AVEVA decision-control layer fitted to large-shipyard process: configuration, impact, evidence, approval and operations feedback connected through customer-owned APIs, governed workflows, policy gates and trusted evidence. This preserves the strengths of specialist tools while positioning AVEVA as the cross-domain decision loop for design change, assurance evidence, production readiness, handover and lifecycle learning.",
        "4. Contribution After Joining",
        "My first contribution would be to align quickly with AVEVA's assigned portfolio, Korea/Japan account priorities, Salesforce pipeline discipline, partner/channel structure, regional marketing activities and collaboration model with sales leadership, Account Managers, Pre-Sales Consultants and Business Development leadership.",
        "I would support Korea/Japan marine opportunities by identifying pain points in design rework, system integration, hull/class configuration, limited digital continuity, compliance workload, production handover, MRO and lifecycle data ownership. These topics can be converted into discovery questions, value propositions, benchmark assumptions, benefit estimates and solution-fit discussions around AVEVA Engineering / Marine, PLM, Engineering Data, Digital Thread and Smart Shipyard.",
        "My contribution would be strongest in customer workshops, executive briefings, Proof of Capability storylines, technical proposal support, competitive positioning, tender-response preparation, feature feedback, issue coordination, implementation handover and customer follow-up. My background with ICDs, test plans, VCRM, evidence packages, ECO/BOM changes, WBS schedules, cost estimates, defect logs and acceptance documents gives me the discipline to turn customer conversations into clear next actions.",
        "I want to help AVEVA grow meaningful Marine Engineering opportunities in Korea and Japan by combining marine-domain credibility, engineering-system understanding, competitive PLM insight and a practical sustainability narrative: connected data that reduces rework, improves productivity, strengthens compliance and helps customers engineer more efficiently."
    )

    $range = $Doc.Range($heading.Range.End, $Doc.Content.End - 1)
    $range.Text = ($paras -join "`r") + "`r"

    foreach ($p in $Doc.Paragraphs) {
        if ($p.Range.Start -gt $heading.Range.Start) {
            $txt = Normalize-Text $p.Range.Text
            if ($txt.Length -gt 0) {
                $p.Range.Font.Name = "Arial"
                $p.Range.Font.Size = [single]10.8
                $p.Range.Font.Bold = [int]0
                $p.Range.ParagraphFormat.SpaceBefore = [single]0
                $p.Range.ParagraphFormat.SpaceAfter = [single]5
                $p.Range.ParagraphFormat.LineSpacingRule = 0
                if ($txt -match "^[1-4]\. ") {
                    $p.Range.Font.Bold = [int]1
                    $p.Range.Font.Size = [single]11
                    $p.Range.ParagraphFormat.SpaceBefore = [single]8
                    $p.Range.ParagraphFormat.SpaceAfter = [single]6
                }
            }
        }
    }
}

$mainDocx = "D:\이력서\AVEVA - Business Development Manager\Joonha_Lee_AVEVA_BDM_Engineering_Marine_Korea_Japan_Seoul_Resume.docx"
$mainPdf = "D:\이력서\AVEVA - Business Development Manager\Joonha_Lee_AVEVA_BDM_Engineering_Marine_Korea_Japan_Seoul_Resume.pdf"
$docx = "D:\이력서\AVEVA - Business Development Manager\Joonha_Lee_AVEVA_BDM_Engineering_Marine_Korea_Japan_Seoul_Resume_Formal_Tone.docx"
$pdf = "D:\이력서\AVEVA - Business Development Manager\Joonha_Lee_AVEVA_BDM_Engineering_Marine_Korea_Japan_Seoul_Resume_Formal_Tone.pdf"
$tempPdf = "C:\Users\namma\.claude\aveva_bdm_resume_work\formal_tone_export.pdf"

if (Test-Path -LiteralPath $tempPdf) { Remove-Item -LiteralPath $tempPdf -Force }

$word = New-Object -ComObject Word.Application
$word.Visible = $false
$doc = $null
try {
    $doc = $word.Documents.Open($docx, $false, $false)
    Replace-SelfIntroductionCompact -Doc $doc
    $doc.Fields.Update() | Out-Null
    $doc.Save()
    $doc.ExportAsFixedFormat($tempPdf, 17)
    $pages = $doc.ComputeStatistics(2)
    $words = $doc.ComputeStatistics(0)
    $doc.Close([ref]$true)
    $doc = $null

    Copy-Item -LiteralPath $tempPdf -Destination $pdf -Force
    Copy-Item -LiteralPath $docx -Destination $mainDocx -Force
    $mainPdfStatus = "SKIPPED"
    try {
        Copy-Item -LiteralPath $pdf -Destination $mainPdf -Force
        $mainPdfStatus = "UPDATED"
    } catch {
        $mainPdfStatus = "LOCKED"
    }

    "UPDATED_DOCX: $docx"
    "UPDATED_PDF: $pdf"
    "MAIN_DOCX: UPDATED"
    "MAIN_PDF: $mainPdfStatus"
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
