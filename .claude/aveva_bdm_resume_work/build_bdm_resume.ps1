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
        [single]$Size = 11,
        [int]$Color = -1
    )
    $cell = $Table.Range.Cells.Item($Index)
    Set-RangeTextKeepMark -Range $cell.Range -Text $Text
    $cell.Range.Font.Name = "Arial"
    $cell.Range.Font.Size = [single]$Size
    $cell.Range.Font.Bold = [int]0
    if ($Bold) { $cell.Range.Font.Bold = [int]1 }
    if ($Color -ge 0) { $cell.Range.Font.Color = [int]$Color }
    $cell.Range.ParagraphFormat.SpaceBefore = [single]0
    $cell.Range.ParagraphFormat.SpaceAfter = [single]0
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
            $hits++
        }
    }
    return $hits
}

function Replace-ParagraphExact {
    param(
        $Doc,
        [string]$OldText,
        [string]$NewText,
        [bool]$Bold = $false
    )
    $needle = Normalize-Text $OldText
    $hits = 0
    foreach ($p in $Doc.Paragraphs) {
        $txt = Normalize-Text $p.Range.Text
        if ($txt -eq $needle) {
            Set-RangeTextKeepMark -Range $p.Range -Text $NewText
            $p.Range.Font.Name = "Arial"
            $p.Range.Font.Size = [single]11
            $p.Range.Font.Bold = [int]0
            if ($Bold) { $p.Range.Font.Bold = [int]1 }
            $hits++
        }
    }
    return $hits
}

function Update-HeaderText {
    param($Doc)
    foreach ($section in $Doc.Sections) {
        foreach ($idx in 1..3) {
            try {
                $header = $section.Headers.Item($idx)
                if ($header.Exists) {
                    foreach ($p in $header.Range.Paragraphs) {
                        $txt = Normalize-Text $p.Range.Text
                        if ($txt.Contains("Joonha Lee")) {
                            Set-RangeTextKeepMark -Range $p.Range -Text "Joonha Lee | AVEVA BDM Engineering (Marine) | Resume"
                            $p.Range.Font.Name = "Arial"
                            $p.Range.Font.Size = [single]10
                            $p.Range.Font.Color = [int]8421504
                            $p.Range.ParagraphFormat.Alignment = 2
                        }
                    }
                }
            } catch {}
        }
    }
}

function Set-TableCells {
    param(
        $Table,
        [object[]]$Rows
    )
    foreach ($row in $Rows) {
        Set-CellText -Table $Table -Index ([int]$row[0]) -Text ([string]$row[1]) -Bold ([bool]$row[2])
    }
}

$source = "D:\이력서\AVEVA - Marine Principal Technical Support & Consultant – PLM SME, Busan\Joonha_Lee_AVEVA_Marine_PLM_SME_Resume.docx"
$outDir = "D:\이력서\AVEVA - Business Development Manager"
$outDocx = Join-Path $outDir "Joonha_Lee_AVEVA_BDM_Engineering_Marine_Korea_Japan_Seoul_Resume.docx"
$outPdf = Join-Path $outDir "Joonha_Lee_AVEVA_BDM_Engineering_Marine_Korea_Japan_Seoul_Resume.pdf"

if (-not (Test-Path -LiteralPath $outDir)) {
    New-Item -ItemType Directory -Path $outDir -Force | Out-Null
}
Copy-Item -LiteralPath $source -Destination $outDocx -Force

$word = New-Object -ComObject Word.Application
$word.Visible = $false
$doc = $null
$log = New-Object System.Collections.Generic.List[string]

try {
    $doc = $word.Documents.Open($outDocx, $false, $false)
    Update-HeaderText -Doc $doc

    $blue = 5570584

    # Table 1: title block.
    Set-CellText -Table $doc.Tables.Item(1) -Index 1 -Text "Resume" -Bold $true -Size ([single]20) -Color $blue
    Set-CellText -Table $doc.Tables.Item(1) -Index 2 -Text "Target: AVEVA - Business Development Manager - Engineering (Marine), Korea / Japan`rApplication Location: Seoul, Korea" -Bold $true -Size ([single]10) -Color $blue

    # Table 2: basic information. Leave the photo cell untouched.
    Set-CellText -Table $doc.Tables.Item(2) -Index 7 -Text "Gunpo-si, Gyeonggi-do, Korea | Application Location: Seoul, Korea"
    Set-CellText -Table $doc.Tables.Item(2) -Index 9 -Text "21+ years"

    # Table 3: core competencies.
    Set-TableCells -Table $doc.Tables.Item(3) -Rows @(
        @(1, "Marine Business Development / Pipeline Support", $true),
        @(2, "Korea shipbuilding and naval-domain credibility for identifying marine customer pain points, qualifying Engineering (Marine) opportunities, and supporting regional sales / account teams with CRM-aware pipeline development.", $false),
        @(3, "Technical Value Selling / Solution Positioning", $true),
        @(4, "Translate design rework, configuration inconsistency, limited digital continuity, compliance inefficiency and lifecycle-data gaps into AVEVA value propositions for PLM, Engineering Data, Smart Shipyard and Digital Thread.", $false),
        @(5, "Marine / Shipbuilding Lifecycle Domain", $true),
        @(6, "ROKN ship/submarine programmes including FFX Batch-II and Jangbogo/KSS-III; requirements, architecture, ICD, FAT/SAT, sea trial, customer acceptance and sustainment evidence.", $false),
        @(7, "Customer Engagement / Executive Communication", $true),
        @(8, "Technical reviews, global customer quality meetings, proposal presentations and issue briefings; able to explain cost, schedule, risk, fleet readiness and lifecycle efficiency in business language.", $false),
        @(9, "Competitive PLM / Engineering Software Insight", $true),
        @(10, "Benchmarked AVEVA E3D/Marine, Unified Engineering, AIM, PI/CONNECT and competitor patterns from Siemens, Dassault, Hexagon, PTC and NAPA to frame open decision-control-plane differentiation.", $false),
        @(11, "AVEVA Values / Sustainability / Trust", $true),
        @(12, "Impact, Aspiration, Curiosity and Trust-aligned working style; practical sustainability story through reduced rework, digital continuity, evidence-based decisions and responsible use of engineering resources.", $false)
    )

    # Table 4: career summary role relevance.
    Set-CellText -Table $doc.Tables.Item(4) -Index 10 -Text "Global satellite-terminal customer engagement; requirements/evidence control, 400+ verification items, issue-to-value framing and release readiness for customer confidence."
    Set-CellText -Table $doc.Tables.Item(4) -Index 15 -Text "RFP/proposal, tender response, cost/schedule planning, technical reviews and customer/partner coordination for defence/aerospace technology opportunities."
    Set-CellText -Table $doc.Tables.Item(4) -Index 20 -Text "15 years Korea naval/shipbuilding customer-facing delivery; Navy/shipyard pain-point discovery, system integration, FAT/SAT, sea trial, acceptance and lifecycle documentation."
    Set-CellText -Table $doc.Tables.Item(4) -Index 25 -Text "Control-board and production issue analysis; technical evidence, certification support and engineering-to-production feedback."
    Set-CellText -Table $doc.Tables.Item(4) -Index 30 -Text "Communication hardware and RF/digital/power interface debugging; early foundation for system-level technical explanation."
    Set-CellText -Table $doc.Tables.Item(4) -Index 35 -Text "Embedded hardware, high-speed interface and prototype-to-production support; evidence-based troubleshooting foundation."

    Replace-ParagraphExact -Doc $doc -OldText "Job-Fit Map: AVEVA Requirements vs. Career Evidence" -NewText "Job-Fit Map: AVEVA Marine BDM Requirements vs. Career Evidence" -Bold $true | Out-Null

    # Table 5: direct JD matching.
    Set-TableCells -Table $doc.Tables.Item(5) -Rows @(
        @(1, "AVEVA BDM Requirement", $true),
        @(2, "Direct Evidence from Career", $true),
        @(3, "Value to AVEVA / Client", $true),
        @(4, "Grow new Marine Business Pipeline and support opportunity qualification", $true),
        @(5, "Supported RFP analysis, technical proposals, customer meetings, issue prioritisation and delivery planning across naval, defence and satellite technology programmes.", $false),
        @(6, "Can help Account Managers convert Korea/Japan marine pain points into qualified AVEVA Engineering opportunities and CRM-aware pipeline inputs.", $false),
        @(7, "Provide deep marine domain, product-market and competitive knowledge", $true),
        @(8, "Worked 15 years in ROKN/shipyard programmes and built an AVEVA Marine PLM benchmark against Siemens, Dassault, Hexagon, PTC and NAPA using the V18 future-strategy deck.", $false),
        @(9, "Gives the sales team domain credibility and a differentiated product narrative for shipbuilding, offshore and naval conversations.", $false),
        @(10, "Build value propositions with benchmarks and benefit estimation", $true),
        @(11, "Managed 400+ verification items, VCRM-style evidence, ECO/BOM change impact, WBS schedules, cost estimates, acceptance criteria and release judgement.", $false),
        @(12, "Can support benefit stories around less rework, faster change impact review, better configuration consistency, lifecycle cost reduction and compliance efficiency.", $false),
        @(13, "Engage executives and speak in business language", $true),
        @(14, "Prepared management reports, proposal presentations, risk/acceptance briefings and customer-facing issue explanations for Navy, shipyard, defence and global partners.", $false),
        @(15, "Can translate technical detail into cost, schedule, risk, fleet readiness, digital transformation and operational-efficiency messages.", $false),
        @(16, "Collaborate with Sales, Pre-Sales, partners and marketing in complex cycles", $true),
        @(17, "Coordinated customers, engineering, quality, production, subcontractors, external agencies and global partners through long-cycle technical delivery and acceptance work.", $false),
        @(18, "Can integrate into strategic opportunity teams with Account Managers, Pre-Sales Consultants, Business Development leadership, partners and Regional Marketing.", $false),
        @(19, "Explain AVEVA technical differentiation in Marine PLM / Engineering", $true),
        @(20, "Developed a future strategy: AVEVA as an open decision control plane linking configuration, impact, evidence, approval and operations feedback above specialist tools.", $false),
        @(21, "Can challenge document-centric workflows and position AVEVA as the engineering-to-operations lifecycle decision loop for marine customers.", $false)
    )

    # Detailed experience: targeted BDM reframing while preserving the original long-form layout.
    $replacements = @(
        @("Managed system integration testing, customer technical communication, quality issue tracking", "Managed system integration testing, global customer technical communication, quality issue tracking, VCRM-based verification control and release-readiness support for the Eutelsat OneWeb LEO satellite terminal and antenna development programme. For the AVEVA BDM Engineering (Marine) role, this is relevant because it shows how to convert customer requirements, technical issues, verification evidence and release risk into business-value discussions around reliability, acceptance readiness and lifecycle traceability."),
        @("Connected customer requirements, field issues, test items, defect history", "Connected customer requirements, field issues, test items, defect history, ECO/BOM changes, verification evidence and release decision criteria into a structured project-control framework. This experience supports AVEVA sales and pre-sales teams by showing how customer pain points can be organised into solution-fit analysis, benefit evidence, opportunity qualification and clear follow-up actions."),
        @("1. System Integration, Verification Traceability and Lifecycle Deliverable Management", "1. Customer Requirement-to-Value Translation and Lifecycle Evidence Management"),
        @("(1) VCRM-Based Requirements and Verification Evidence Control", "(1) Customer Requirements, Evidence and Business Value Control"),
        @("(2) Integration Verification Based on 400+ Design Verification Items", "(2) Benchmark and Benefit Evidence from 400+ Verification Items"),
        @("(3) Product Issue Tracking and Regression Verification Control", "(3) Customer Pain Point Classification and Conversion to Action"),
        @("2. Global Customer Technical Support and Product Feedback Coordination", "2. Global Customer Engagement and Sales / Pre-Sales Support"),
        @("(1) Customer Quality Meetings and Technical Reviews", "(1) Customer Meetings, Technical Discovery and Follow-Up"),
        @("(2) Issue Alignment Between Customer and Product Teams", "(2) Opportunity-Grade Issue Alignment Between Customer and Product Teams"),
        @("3. Production Transition, Release Readiness and Field Technical Support", "3. Product Readiness, Release Risk and Commercial Impact Support"),
        @("Managed defence and aerospace electronics projects, including K2 sight test equipment", "Managed defence and aerospace electronics projects, including K2 sight test equipment and civil aircraft diagnostic system development, from RFP/proposal through testing and delivery. This experience is relevant to AVEVA BDM work because it required tender interpretation, value-oriented technical explanation, schedule/cost planning, partner coordination and customer-facing delivery discipline."),
        @("Led RFP analysis, technical proposal preparation, project execution planning", "Led RFP analysis, technical proposal preparation, project execution planning, verification readiness review, deliverable control and coordination with customers, partners and external test agencies. The same pattern can support AVEVA opportunity qualification, customer discovery, proposal support, benchmark explanation and benefit estimation."),
        @("1. Proposal, Contract and Full Project Execution Management", "1. Tender, Proposal and Value-Based Opportunity Support"),
        @("(2) RFP Analysis and Technical Proposal Preparation", "(2) RFP Analysis, Solution Positioning and Technical Proposal Preparation"),
        @("Delivered ROK Navy ship and submarine communication systems", "Delivered ROK Navy ship and submarine communication systems, shipboard broadcasting systems, internal networks, CCTV/VOD and C4I-linked systems from requirements analysis through design, installation, commissioning, acceptance and after-delivery support. This is the strongest foundation for AVEVA Marine BDM because it gives practical understanding of Korean shipyard workflows, naval customer language, system integration pain points and lifecycle documentation needs."),
        @("Managed project execution, marine system integration, technical documentation", "Managed project execution, marine system integration, technical documentation, FAT/SAT, sea trial support, customer acceptance, field issue response and coordination with the Navy, shipyards, inspectors, subcontractors, production, quality and engineering teams. This experience supports marine customer discovery, executive risk conversation and AVEVA value positioning around digital thread, configuration control, change impact and evidence-based approval."),
        @("1. End-to-End PM for Naval and Shipbuilding SI Projects", "1. End-to-End Marine Customer Engagement for Naval and Shipbuilding SI Projects"),
        @("(2) Requirements Analysis and Proposal / Execution Planning", "(2) Requirements Analysis, Customer Pain-Point Discovery and Proposal Planning"),
        @("2. Marine System Integration and Lifecycle Documentation Control", "2. Marine Engineering Data, System Integration and Lifecycle Documentation Control"),
        @("(3) Shipboard Technical Data and Customer Acceptance Control", "(3) Shipboard Technical Data, Customer Acceptance and Lifecycle Value Control")
    )
    foreach ($item in $replacements) {
        Replace-ParagraphContains -Doc $doc -Fragment $item[0] -NewText $item[1] -Bold ($item[1] -match "^\d\.|^\(") | Out-Null
    }

    # Table 12: toolkit.
    Set-CellText -Table $doc.Tables.Item(12) -Index 2 -Text "Certified NMEA 2000 Networking; Defence PMBOK training; Machine Learning with Python; Python for Data Science, AI & Development; AVEVA Marine PLM / Engineering Data / Digital Thread preparation."
    Set-CellText -Table $doc.Tables.Item(12) -Index 4 -Text "IELTS 5.0; English language study in Perth, Australia; global customer technical meetings and English technical communication with Eutelsat OneWeb and international partners; ready to support Korea/Japan regional opportunities in English."
    Set-CellText -Table $doc.Tables.Item(12) -Index 6 -Text "AVEVA E3D/Marine concepts, Unified Engineering, AIM, PI/CONNECT positioning, PLM, Digital Thread, Digital Twin, Smart Shipyard, ERP/MES/PLM boundaries, EPLAN concept, ICD, VCRM, ECO/BOM, WBS, cost estimate, defect log, evidence package, Wireshark, iperf, Linux, VMware and RF measurement equipment."

    # Self-introduction section.
    Replace-ParagraphContains -Doc $doc -Fragment "I understand that the AVEVA Marine Principal Technical Support" -NewText "I understand that the AVEVA Business Development Manager - Engineering (Marine), Korea / Japan role is not a pure sales title and not a pure engineering title. It is a senior customer-facing business-development role that must connect marine industry domain knowledge, AVEVA Engineering / Marine product value, competitive differentiation, opportunity qualification, regional sales support and executive-level customer conversations. I am applying for the Seoul, Korea location." | Out-Null
    Replace-ParagraphContains -Doc $doc -Fragment "My motivation for this role comes from more than 21 years" -NewText "My motivation comes from more than 21 years of experience in high-reliability electronics, naval shipbuilding projects, defence systems, satellite communication systems, customer technical support, proposal work, verification evidence and cross-functional delivery. During 15 years at Daeyang Electric, I worked on ROK Navy ship and submarine programmes such as FFX Batch-II and Jangbogo/KSS-III, coordinating with Navy users, shipyards, inspectors, subcontractors, engineering, production, quality and project teams." | Out-Null
    Replace-ParagraphContains -Doc $doc -Fragment "Through these shipbuilding and naval projects" -NewText "Through these projects, I learned that marine customers buy more than software features. They need lower rework, clearer configuration control, faster change-impact review, reliable evidence for approval, better integration between engineering and production, and a credible digital thread from design through build, sustainment and operations. This is exactly the business language behind AVEVA's Engineering (Marine) value proposition." | Out-Null
    Replace-ParagraphContains -Doc $doc -Fragment "I am interested in AVEVA because its marine solutions" -NewText "I am interested in AVEVA because AVEVA can help marine customers move from document-centric execution to trusted, connected engineering data. I also understand AVEVA's sustainability direction: industrial software should help customers use resources more responsibly by improving efficiency, reducing waste/rework and enabling better decisions with trusted data. I understand AVEVA values as Impact, Aspiration, Curiosity, and Trust. These are close to my working style. In complex projects, I try to make practical impact, understand the customer goal, ask why before deciding, and build trust through clear evidence and follow-up." | Out-Null

    Replace-ParagraphContains -Doc $doc -Fragment "My main strength is the ability to connect technical understanding" -NewText "My main strength is the ability to bridge marine engineering reality and business development execution. I can listen to shipyard engineers, identify the real pain behind a technical symptom, translate it into customer value, and help sales / pre-sales teams structure the next action: discovery question, value proposition, benchmark, PoC scope, risk item, owner, or qualified opportunity." | Out-Null
    Replace-ParagraphContains -Doc $doc -Fragment "At Daeyang Electric, I managed and supported naval communication" -NewText "At Daeyang Electric, I managed and supported naval communication and shipboard systems from requirements analysis through proposal, design, installation, FAT, SAT, sea trial, customer acceptance and after-delivery support. I converted Navy and shipyard requirements into system architecture, equipment configuration, ICDs, test plans, test procedures, acceptance criteria and close-out documents. This experience gives me credible working knowledge for Korean marine customer engagement." | Out-Null
    Replace-ParagraphContains -Doc $doc -Fragment "At Intellian Technologies, I further strengthened this capability" -NewText "At Intellian Technologies, I strengthened this capability in a global customer environment by managing requirements, field issues, defect history, ECO/BOM changes, verification evidence and release judgement for a LEO satellite-terminal programme. I also operated more than 400 design verification checklist items across RF/antenna, control electronics, network interface, environmental verification and production verification areas. This is useful for AVEVA value selling because customer trust is built by evidence, not claims." | Out-Null
    Replace-ParagraphContains -Doc $doc -Fragment "Although my background is not limited to PLM software development itself" -NewText "Although I have not held a formal Business Development Manager title or direct Salesforce quota ownership, my background fits the technical-value-selling side of this role: customer discovery, issue classification, proposal support, benefit explanation, risk reduction, lifecycle data thinking, and collaboration with account managers, pre-sales consultants, technical teams, partners and leadership. I will describe pipeline work accurately as CRM-aware pipeline support and opportunity qualification support." | Out-Null

    Replace-ParagraphExact -Doc $doc -OldText "3. Marine PLM, Customer Support and Product Requirement Communication" -NewText "3. AVEVA Marine PLM Analysis, Competitive Differentiation and Future Strategy" -Bold $true | Out-Null
    Replace-ParagraphContains -Doc $doc -Fragment "My approach to technical support is structured" -NewText "When I previously prepared for the AVEVA Marine PLM consultant role, I did not stop at interview preparation. I analysed AVEVA Marine / PLM from the viewpoint of current product position, competitor strengths, customer pain points and future strategy. Based on the V18 meeting deck, my conclusion is that AVEVA should not compete only as another closed PLM backbone. AVEVA can win by becoming the engineering-to-operations decision control plane above specialist tools." | Out-Null
    Replace-ParagraphContains -Doc $doc -Fragment "In naval shipbuilding projects, I worked directly with Navy users" -NewText "My comparison view is practical: Siemens Teamcenter is strong in configuration, BOM and change governance; Dassault 3DEXPERIENCE is strong in collaboration and virtual-twin positioning; Hexagon is strong in asset lifecycle and operations context; PTC is strong in requirements, ALM and traceability patterns; NAPA is trusted for naval architecture and stability calculation. AVEVA's advantage is to connect engineering truth, lifecycle context and operations feedback through E3D/Marine, Unified Engineering, AIM, PI/CONNECT and open integration." | Out-Null
    Replace-ParagraphContains -Doc $doc -Fragment "At Intellian Technologies, I also operated regular quality meetings" -NewText "The future strategy I would explain to marine customers is an open AVEVA decision control plane: configuration, impact, evidence, approval and operations feedback connected through customer-owned APIs, governed workflows, policy gates and trusted evidence. This approach keeps specialist tools where they are strong, while allowing AVEVA to own the cross-domain decision loop for design change, assurance evidence, production readiness, handover and lifecycle learning." | Out-Null
    Replace-ParagraphContains -Doc $doc -Fragment "For AVEVA, I can contribute by collecting customer expectations" -NewText "This is why I strongly want to work for AVEVA. The company is positioned to turn marine engineering data into business outcomes: lower design rework, faster change impact, stronger configuration consistency, better compliance evidence, improved production efficiency, reduced lifecycle risk and more sustainable use of industrial resources. I can help make this message practical for Korea/Japan customers through customer discovery, value proposition development, PoC planning and evidence-based follow-up." | Out-Null

    Replace-ParagraphContains -Doc $doc -Fragment "If selected, I would aim to contribute as a practical Marine PLM technical consultant" -NewText "If selected, I would contribute as a Marine Engineering Business Development / Technical Value Selling professional who understands both shipbuilding project reality and AVEVA's industrial software opportunity. My first priority would be to understand AVEVA's assigned product portfolio, Korea/Japan account priorities, sales process, Salesforce CRM pipeline discipline, partner/channel structure, regional marketing campaigns and internal collaboration flow with sales leadership, account managers, pre-sales consultants and business development leadership." | Out-Null
    Replace-ParagraphContains -Doc $doc -Fragment "I would then focus on supporting Korean shipyard customers" -NewText "I would then support Korea/Japan marine opportunities by identifying customer pain points in design rework, system integration, hull-class configuration, limited digital continuity, regulatory/compliance effort, production handover, MRO and lifecycle data ownership. I can translate these points into discovery questions, value propositions, benchmark assumptions, benefit estimates and solution-fit discussions around AVEVA Engineering / Marine, PLM, Engineering Data, Digital Thread and Smart Shipyard." | Out-Null
    Replace-ParagraphContains -Doc $doc -Fragment "I can also contribute to customer workshops" -NewText "I can contribute to customer workshops, executive briefings, Proof of Capability storylines, technical proposal support, competitive positioning, tender-response preparation, feature feedback, issue coordination, implementation handover and customer follow-up. My experience with ICDs, test plans, VCRM, evidence packages, ECO/BOM changes, WBS schedules, cost estimates, defect logs and acceptance documents will help me organise customer conversations into clear next actions." | Out-Null
    Replace-ParagraphContains -Doc $doc -Fragment "I understand that AVEVA’s role in the marine industry" -NewText "I understand that AVEVA's Commercial team needs people who can deeply understand customer needs and deliver tailored solutions. I can add value by combining marine domain credibility, engineering-system understanding, customer trust building, competitive PLM insight and a practical sustainability narrative: use connected data to reduce rework, improve productivity, support compliance and help customers engineer more efficiently." | Out-Null
    Replace-ParagraphContains -Doc $doc -Fragment "I am prepared to deepen my knowledge of AVEVA Marine PLM" -NewText "I am prepared to deepen my knowledge of AVEVA E3D/Marine, Unified Engineering, AIM, PI/CONNECT, CONNECT AI themes, AVEVA's partner ecosystem and Korea/Japan marine market priorities with discipline and urgency. I will bring Impact, Aspiration, Curiosity and Trust into daily work by making practical impact, aiming for customer-visible outcomes, asking why before deciding, and building trust through clear evidence and follow-up. With this approach, I believe I can help AVEVA grow marine business pipeline and win meaningful Engineering (Marine) opportunities in Korea and Japan." | Out-Null

    $doc.Fields.Update() | Out-Null
    $doc.Save()
    $doc.ExportAsFixedFormat($outPdf, 17)
    $pages = $doc.ComputeStatistics(2)
    $words = $doc.ComputeStatistics(0)
    $log.Add("DOCX: $outDocx")
    $log.Add("PDF: $outPdf")
    $log.Add("PAGES: $pages")
    $log.Add("WORDS: $words")
    $doc.Close([ref]$true)
    $doc = $null
} finally {
    if ($doc -ne $null) {
        try { $doc.Close([ref]$false) } catch {}
    }
    if ($word -ne $null -and $word.Documents.Count -eq 0) {
        try { $word.Quit() } catch {}
    }
}

$report = "C:\Users\namma\.claude\aveva_bdm_resume_work\build_bdm_resume_report.txt"
[System.IO.File]::WriteAllLines($report, $log, [System.Text.Encoding]::UTF8)
$log
