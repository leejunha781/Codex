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

function Set-CellText {
    param(
        $Table,
        [int]$Index,
        [string]$Text,
        [bool]$Bold = $false,
        [single]$Size = 10.0
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

function Find-ParagraphExact {
    param($Doc, [string]$Text)
    $needle = Normalize-Text $Text
    foreach ($p in $Doc.Paragraphs) {
        if ((Normalize-Text $p.Range.Text) -eq $needle) { return $p }
    }
    return $null
}

function Replace-ParagraphContains {
    param(
        $Doc,
        [string]$Fragment,
        [string]$NewText,
        [single]$Size = 11.0,
        [bool]$Bold = $false
    )
    $needle = Normalize-Text $Fragment
    $hits = 0
    foreach ($p in $Doc.Paragraphs) {
        $txt = Normalize-Text $p.Range.Text
        if ($txt.Length -gt 0 -and $txt.Contains($needle)) {
            Set-RangeTextKeepMark -Range $p.Range -Text $NewText
            $p.Range.Font.Name = "Arial"
            $p.Range.Font.Size = [single]$Size
            $p.Range.Font.Bold = [int]0
            if ($Bold) { $p.Range.Font.Bold = [int]1 }
            $p.Range.ParagraphFormat.SpaceBefore = [single]0
            $p.Range.ParagraphFormat.SpaceAfter = [single]8
            $p.Range.ParagraphFormat.LineSpacingRule = 0
            $hits++
        }
    }
    return $hits
}

function Replace-SelfIntroduction {
    param($Doc)
    $heading = Find-ParagraphExact -Doc $Doc -Text "Self-Introduction"
    if ($null -eq $heading) { throw "Self-Introduction heading not found" }

    $newText = @"

1. Motivation for Application

I am applying for AVEVA's Business Development Manager - Engineering (Marine), Korea / Japan role in Seoul because it sits at the intersection of marine engineering, customer value and industrial software. My career has been built in environments where customer requirements, technical risk, verification evidence and delivery commitments must be converted into practical decisions.

For more than 21 years, I have worked across naval shipbuilding, defence electronics and satellite communication systems. The strongest foundation is my 15 years at Daeyang Electric, where I supported ROK Navy ship and submarine programmes including FFX Batch-II and Jangbogo/KSS-III, coordinating with Navy users, shipyards, inspectors, subcontractors, engineering, production, quality and project teams.

Marine customers do not buy software functions alone. They need lower rework, stronger configuration control, faster change-impact review, reliable approval evidence and a trusted digital thread from design through build, handover and operations. That is the business value I see in AVEVA Engineering / Marine.

AVEVA's values of Impact, Aspiration, Curiosity and Trust are close to my working style. I aim for practical impact, ask why before deciding, and build trust through evidence, clear follow-up and accountable communication. I also respect AVEVA's sustainability direction: better engineering data should reduce waste, improve productivity and help customers use industrial resources more responsibly.

2. Fit for the Role

My value is the ability to translate shipbuilding reality into business-development action. I am comfortable listening to engineers, identifying the real issue behind a technical symptom, and turning it into a discovery question, value proposition, benchmark assumption, PoC scope, risk item or qualified opportunity.

At Daeyang Electric, I converted Navy and shipyard requirements into system architecture, equipment configuration, ICDs, test plans, procedures, acceptance criteria and close-out documents. That experience connects naturally with the HD Hyundai-standard process: requirements/VCRM, E3D/Marine/Draw, MTO/eBOM, baseline/effectivity, ECR/ECO and handover evidence.

At Intellian Technologies, I worked in a global customer environment, managing requirements, field issues, defect history, ECO/BOM changes, verification evidence and release judgement for a LEO satellite-terminal programme. Managing more than 400 verification checklist items reinforced one lesson: customer trust is earned through evidence, not claims.

Commercially, this background fits the technical value-selling side of the role: customer discovery, proposal support, benefit explanation, risk reduction, lifecycle data thinking and cross-functional follow-up. I would support Account Managers and Pre-Sales Consultants with CRM-aware opportunity qualification, discovery evidence, value-proposition inputs and disciplined next actions.

3. AVEVA Marine PLM Perspective

While preparing for the AVEVA Marine PLM consultant role, I analysed AVEVA Marine / PLM from the perspectives of product position, competitor strengths, customer pain points and future strategy. My conclusion from the V18 meeting deck was clear: AVEVA should not compete only as a closed PLM backbone. Its stronger position is an engineering-to-operations decision layer above specialist tools.

My comparison view is practical. Siemens Teamcenter is strong in configuration, BOM and change governance; Dassault 3DEXPERIENCE in collaboration and virtual-twin positioning; Hexagon in asset lifecycle and operations; PTC in requirements, ALM and traceability; and NAPA in naval architecture and stability calculation. AVEVA's opportunity is to connect engineering truth, lifecycle context and operations feedback through E3D/Marine, Unified Engineering, AIM, PI/CONNECT and open integration.

The customer message I would carry is an open AVEVA decision-control layer fitted to large-shipyard process: configuration, impact, evidence, approval and operations feedback connected through customer-owned APIs, governed workflows, policy gates and trusted evidence. This preserves the strengths of specialist tools while positioning AVEVA as the cross-domain decision loop for design change, assurance evidence, production readiness, handover and lifecycle learning.

4. Contribution After Joining

My first contribution would be to align quickly with AVEVA's assigned portfolio, Korea/Japan account priorities, Salesforce pipeline discipline, partner/channel structure, regional marketing activities and collaboration model with sales leadership, Account Managers, Pre-Sales Consultants and Business Development leadership.

I would support Korea/Japan marine opportunities by identifying pain points in design rework, system integration, hull/class configuration, limited digital continuity, compliance workload, production handover, MRO and lifecycle data ownership. These topics can be converted into discovery questions, value propositions, benchmark assumptions, benefit estimates and solution-fit discussions around AVEVA Engineering / Marine, PLM, Engineering Data, Digital Thread and Smart Shipyard.

My contribution would be strongest in customer workshops, executive briefings, Proof of Capability storylines, technical proposal support, competitive positioning, tender-response preparation, feature feedback, issue coordination, implementation handover and customer follow-up. My background with ICDs, test plans, VCRM, evidence packages, ECO/BOM changes, WBS schedules, cost estimates, defect logs and acceptance documents gives me the discipline to turn customer conversations into clear next actions.

I want to help AVEVA grow meaningful Marine Engineering opportunities in Korea and Japan by combining marine-domain credibility, engineering-system understanding, competitive PLM insight and a practical sustainability narrative: connected data that reduces rework, improves productivity, strengthens compliance and helps customers engineer more efficiently.
"@

    $start = $heading.Range.End
    $end = $Doc.Content.End - 1
    $range = $Doc.Range($start, $end)
    $range.Text = $newText

    foreach ($p in $Doc.Paragraphs) {
        if ($p.Range.Start -gt $heading.Range.Start) {
            $txt = Normalize-Text $p.Range.Text
            if ($txt.Length -gt 0) {
                $p.Range.Font.Name = "Arial"
                $p.Range.Font.Size = [single]11
                $p.Range.Font.Bold = [int]0
                $p.Range.ParagraphFormat.SpaceBefore = [single]0
                $p.Range.ParagraphFormat.SpaceAfter = [single]8
                $p.Range.ParagraphFormat.LineSpacingRule = 0
                if ($txt -match "^[1-4]\. ") {
                    $p.Range.Font.Bold = [int]1
                    $p.Range.ParagraphFormat.SpaceBefore = [single]8
                }
            }
        }
    }
}

$docx = "D:\이력서\AVEVA - Business Development Manager\Joonha_Lee_AVEVA_BDM_Engineering_Marine_Korea_Japan_Seoul_Resume.docx"
$pdf = "D:\이력서\AVEVA - Business Development Manager\Joonha_Lee_AVEVA_BDM_Engineering_Marine_Korea_Japan_Seoul_Resume.pdf"

$word = New-Object -ComObject Word.Application
$word.Visible = $false
$doc = $null
try {
    $doc = $word.Documents.Open($docx, $false, $false)

    # Core Competencies: tighter, more formal value statements.
    $core = $doc.Tables.Item(3)
    Set-CellText -Table $core -Index 2 -Text "Shipbuilding and naval-domain credibility for discovering customer pain points, qualifying Marine Engineering opportunities and supporting regional account teams." -Size ([single]10)
    Set-CellText -Table $core -Index 4 -Text "Value-selling language for design rework, configuration inconsistency, limited digital continuity, compliance workload and lifecycle-data gaps." -Size ([single]10)
    Set-CellText -Table $core -Index 6 -Text "Practical understanding of ROKN programmes and the HD Hyundai-standard lifecycle: requirements/VCRM, basic/detail design, E3D/Marine/Draw, MTO/eBOM, baseline/effectivity, ECR/ECO and handover evidence." -Size ([single]10)
    Set-CellText -Table $core -Index 8 -Text "Customer meetings, technical reviews, proposal presentations and executive briefings translated into cost, schedule, risk, fleet readiness and lifecycle efficiency." -Size ([single]10)
    Set-CellText -Table $core -Index 10 -Text "Working comparison of AVEVA E3D/Marine, Unified Engineering, AIM and PI/CONNECT against Siemens, Dassault, Hexagon, PTC and NAPA." -Size ([single]10)
    Set-CellText -Table $core -Index 12 -Text "Working style aligned with Impact, Aspiration, Curiosity and Trust; sustainability narrative built on reduced rework, digital continuity, evidence-based decisions and responsible use of engineering resources." -Size ([single]10)

    # Career Summary: concise, role-relevant evidence without explanatory filler.
    $career = $doc.Tables.Item(4)
    Set-CellText -Table $career -Index 10 -Text "Led customer-facing SIT, verification governance and release-readiness control for a global LEO terminal programme; managed 400+ verification items." -Size ([single]10)
    Set-CellText -Table $career -Index 15 -Text "Managed defence/aerospace RFPs, tender responses, cost/schedule planning, technical reviews and customer/partner coordination." -Size ([single]10)
    Set-CellText -Table $career -Index 20 -Text "Built 15 years of naval/shipbuilding delivery credibility across requirements, design, FAT/SAT, sea trial and handover; mapped shipyard pain points to HD Hyundai-standard lifecycle controls." -Size ([single]10)
    Set-CellText -Table $career -Index 25 -Text "Resolved control-board and production issues through evidence-based analysis, certification support and engineering-to-production feedback." -Size ([single]10)
    Set-CellText -Table $career -Index 30 -Text "Developed RF/digital/power interface troubleshooting capability for communication hardware and system-level issue analysis." -Size ([single]10)
    Set-CellText -Table $career -Index 35 -Text "Supported embedded hardware, high-speed interface testing and prototype-to-production stabilization through disciplined failure analysis." -Size ([single]10)

    # Job-Fit Map: remove "Can help" tone and use direct business value.
    $fit = $doc.Tables.Item(5)
    Set-CellText -Table $fit -Index 4 -Text "Grow Marine pipeline and qualify opportunities" -Bold $true -Size ([single]9.7)
    Set-CellText -Table $fit -Index 5 -Text "Supported RFP analysis, technical proposals, customer meetings and issue prioritisation across naval, defence and satellite programmes." -Size ([single]9.7)
    Set-CellText -Table $fit -Index 6 -Text "Converts Korea/Japan marine pain points into qualified Engineering opportunities and CRM-ready next actions." -Size ([single]9.7)
    Set-CellText -Table $fit -Index 7 -Text "Bring marine domain and competitive insight" -Bold $true -Size ([single]9.7)
    Set-CellText -Table $fit -Index 8 -Text "15 years in ROK Navy and shipyard programmes; practical view of requirements/VCRM, E3D/Marine/Draw, MTO/eBOM, baseline/effectivity, ECR/ECO and handover evidence; benchmarked AVEVA against Siemens, Dassault, Hexagon, PTC and NAPA." -Size ([single]9.7)
    Set-CellText -Table $fit -Index 9 -Text "Creates credible discovery with shipyard, offshore and naval customers, grounded in real lifecycle process." -Size ([single]9.7)
    Set-CellText -Table $fit -Index 10 -Text "Build value propositions and benefit cases" -Bold $true -Size ([single]9.7)
    Set-CellText -Table $fit -Index 11 -Text "Managed 400+ verification items, VCRM-style evidence, ECO/BOM impact, WBS schedules, cost estimates and acceptance criteria." -Size ([single]9.7)
    Set-CellText -Table $fit -Index 12 -Text "Supports benefit cases for lower rework, faster change-impact review, stronger configuration consistency and compliance efficiency." -Size ([single]9.7)
    Set-CellText -Table $fit -Index 13 -Text "Engage executives in business language" -Bold $true -Size ([single]9.7)
    Set-CellText -Table $fit -Index 14 -Text "Prepared management reports, proposal presentations and risk/acceptance briefings for Navy, shipyard, defence and global customer stakeholders." -Size ([single]9.7)
    Set-CellText -Table $fit -Index 15 -Text "Translates technical detail into cost, schedule, risk, readiness and sustainability value." -Size ([single]9.7)
    Set-CellText -Table $fit -Index 16 -Text "Coordinate sales, pre-sales, partners and delivery teams" -Bold $true -Size ([single]9.7)
    Set-CellText -Table $fit -Index 17 -Text "Coordinated engineering, quality, production, subcontractors, agencies and global partners through long-cycle delivery and acceptance work." -Size ([single]9.7)
    Set-CellText -Table $fit -Index 18 -Text "Fits regional account teams that require disciplined follow-up across sales, pre-sales, partners and customer functions." -Size ([single]9.7)
    Set-CellText -Table $fit -Index 19 -Text "Position AVEVA Marine PLM / Engineering differentiation" -Bold $true -Size ([single]9.7)
    Set-CellText -Table $fit -Index 20 -Text "Defined an AVEVA narrative around an open decision layer linking configuration, impact, evidence, approval and operations feedback." -Size ([single]9.7)
    Set-CellText -Table $fit -Index 21 -Text "Positions AVEVA above document-centric workflows as the engineering-to-operations decision loop for marine customers." -Size ([single]9.7)

    # Detailed experience: remove resume-explanation tone from the most visible paragraphs.
    Replace-ParagraphContains -Doc $doc -Fragment "Managed system integration testing, global customer technical communication, quality issue tracking" -NewText "Managed system integration testing, customer technical communication, quality-issue governance and release-readiness control for the Eutelsat OneWeb LEO satellite terminal and antenna programme. Converted requirements, defects, ECO/BOM changes and verification evidence into clear release decisions for engineering, quality, production and customer stakeholders." | Out-Null
    Replace-ParagraphContains -Doc $doc -Fragment "Connected customer requirements, field issues, test items, defect history" -NewText "Established a project-control discipline linking customer requirements, field issues, defect history, test evidence and approval criteria. The same discipline applies to Marine BDM work: turning customer pain points into solution fit, benefit evidence and qualified next actions." | Out-Null
    Replace-ParagraphContains -Doc $doc -Fragment "Managed defence and aerospace electronics projects, including K2 sight test equipment" -NewText "Managed defence and aerospace electronics projects, including K2 sight test equipment and civil aircraft diagnostic systems, from RFP/proposal through development, testing, customer inspection and delivery. Led tender interpretation, technical proposal writing, cost/schedule planning, partner coordination and customer-facing reviews." | Out-Null
    Replace-ParagraphContains -Doc $doc -Fragment "Led RFP analysis, technical proposal preparation, project execution planning" -NewText "Controlled RFP analysis, execution planning, verification readiness and deliverables across customers, partners and external test agencies. Built a proposal-to-delivery rhythm suitable for opportunity qualification, benchmark discussion and value-based customer follow-up." | Out-Null
    Replace-ParagraphContains -Doc $doc -Fragment "Delivered ROK Navy ship and submarine communication systems, shipboard broadcasting systems" -NewText "Delivered ROK Navy ship and submarine communication systems, shipboard broadcasting systems, internal networks, CCTV/VOD and C4I-linked systems from requirements analysis through design, installation, commissioning, acceptance and after-delivery support. Developed first-hand understanding of Korean shipyard workflows, naval customer language, system integration constraints and lifecycle documentation." | Out-Null
    Replace-ParagraphContains -Doc $doc -Fragment "Managed project execution, marine system integration, technical documentation" -NewText "Managed marine system integration, technical documentation, FAT/SAT, sea trial support, customer acceptance, field issue response and coordination with Navy, shipyards, inspectors, subcontractors, production, quality and engineering teams. Mapped this delivery experience to large-shipyard lifecycle controls: requirements/VCRM, basic/detail design, E3D/Marine/Draw, MTO/eBOM, hull/block baseline, ECR/ECO, production release and handover evidence." | Out-Null
    Replace-ParagraphContains -Doc $doc -Fragment "Built practical experience in board-level troubleshooting, product stabilisation" -NewText "Established a disciplined approach to board-level troubleshooting, product stabilisation, technical documentation and issue feedback between engineering, quality and production teams, with emphasis on interface clarification, verification evidence and engineering-to-production handover." | Out-Null
    Replace-ParagraphContains -Doc $doc -Fragment "Strengthened practical understanding of RF signal paths, digital control interfaces" -NewText "Strengthened practical understanding of RF signal paths, digital control interfaces, power behaviour, measurement-based debugging and communication equipment reliability, connecting technical data, interface conditions, verification results and field issues across the product lifecycle." | Out-Null
    Replace-ParagraphContains -Doc $doc -Fragment "Developed a practical foundation in circuit behaviour, power integrity" -NewText "Developed a practical foundation in circuit behaviour, power integrity, signal-path analysis, firmware interaction, interface stability and manufacturability. This foundation later supported work in defence electronics, naval shipboard systems and satellite communication systems, where lifecycle documentation, interface traceability and customer acceptance evidence are critical." | Out-Null

    Replace-SelfIntroduction -Doc $doc

    foreach ($tb in $doc.Tables) {
        foreach ($cell in $tb.Range.Cells) {
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
