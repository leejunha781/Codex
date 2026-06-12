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

function Replace-SelfIntroduction {
    param($Doc)

    $heading = Find-ParagraphExact -Doc $Doc -Text "Self-Introduction"
    if ($null -eq $heading) { throw "Self-Introduction heading not found" }

    $paras = @(
        "1. Motivation and AVEVA Fit",
        "I am applying for the Business Development Manager - Engineering (Marine), Korea / Japan role in Seoul because it requires the exact combination I have built throughout my career: marine engineering credibility, customer-facing technical judgement, and the ability to translate complex project issues into business value. This role is not only about selling software; it is about understanding how shipyards, owners, engineering teams and executives make decisions when schedule, quality, compliance and lifecycle risk are all connected.",
        "For more than 21 years, I have worked in high-reliability engineering environments covering naval shipbuilding, defence electronics, satellite communication systems and customer technical delivery. My strongest foundation is the 15 years I spent at Daeyang Electric, supporting ROK Navy ship and submarine programmes including FFX Batch-II and Jangbogo/KSS-III. In those projects, requirements, interfaces, test evidence, FAT/SAT, sea trials, acceptance records and after-delivery support had to be managed with discipline, because customer trust depended on clear evidence and reliable follow-up.",
        "That experience is the reason AVEVA is highly attractive to me. Marine customers do not purchase software functions in isolation. They are looking for lower design rework, stronger configuration control, faster change-impact review, reliable approval evidence, better production handover and a trusted digital thread from engineering through build, operation and lifecycle learning. AVEVA Engineering / Marine is positioned to turn those engineering problems into measurable business outcomes.",
        "I understand AVEVA values Impact, Aspiration, Curiosity and Trust. These values are close to my working style. In complex projects, I try to make practical impact, understand the customer goal, ask why before deciding, and build trust through clear evidence and follow-up. I also respect AVEVA's sustainability direction: connected engineering data should reduce waste, improve productivity, support better decisions and help customers use industrial resources more responsibly.",
        "2. Marine Customer and Lifecycle Expertise",
        "My core strength is the ability to stand between the customer, engineering teams and business stakeholders, then convert technical reality into a clear action plan. In shipbuilding and defence projects, I learned that the real issue is often not a single defect or requirement. It is the lack of a shared view across configuration, interface, schedule, test evidence, approval status and downstream impact. This is exactly where a strong Marine BDM must create value.",
        "At Daeyang Electric, I converted Navy and shipyard requirements into system architecture, equipment configuration, ICDs, test plans, test procedures, acceptance criteria and close-out documents. I also supported installation, commissioning, FAT, SAT, sea trial and customer acceptance activities with Navy users, shipyards, inspectors, subcontractors, production, quality and engineering teams. This gives me practical credibility when discussing real shipyard pain points, not only software concepts.",
        "I can connect that experience to the HD Hyundai-standard shipbuilding lifecycle: requirements/VCRM, basic and detail design, E3D/Marine/Draw, MTO/eBOM, baseline and effectivity control, ECR/ECO, production release and handover evidence. This process view is important for AVEVA because many customer conversations begin with symptoms such as document inconsistency, delayed change review, unclear ownership of engineering data, weak handover evidence or fragmented digital continuity.",
        "At Intellian Technologies, I expanded this discipline in a global customer environment by managing requirements, field issues, defect history, ECO/BOM changes, verification evidence and release judgement for a LEO satellite-terminal programme. More than 400 verification checklist items had to be controlled across RF/antenna, control electronics, network interface, environmental verification and production verification. That experience reinforced a principle I would bring to AVEVA customer engagement: trust is built through evidence, not claims.",
        "Commercially, this background fits the technical value-selling side of the BDM role. I can support Account Managers and Pre-Sales Consultants by structuring discovery questions, clarifying customer pain points, preparing value-proposition inputs, shaping Proof of Capability storylines and turning technical discussion into CRM-ready next actions. My contribution is strongest where customer needs must be interpreted accurately before a solution can be positioned credibly.",
        "3. AVEVA Marine PLM Perspective",
        "When I prepared for the AVEVA Marine PLM consultant role, I did not treat it as short-term interview preparation. I studied AVEVA Marine / PLM from the viewpoints of product position, competitor strengths, customer pain points and future strategy, and developed my own view through the V18 meeting deck. That preparation strengthened my motivation to work for AVEVA because I could see a clear strategic opportunity in the marine industry.",
        "My practical comparison is as follows. Siemens Teamcenter is strong in configuration, BOM and change governance. Dassault 3DEXPERIENCE is strong in collaboration and virtual-twin positioning. Hexagon has strength in asset lifecycle and operations context. PTC is strong in requirements, ALM and traceability patterns. NAPA is trusted in naval architecture and stability calculation. AVEVA's opportunity is different: it can connect engineering truth, lifecycle context and operational feedback through E3D/Marine, Unified Engineering, AIM, PI/CONNECT and open integration.",
        "For Korea and Japan marine customers, I would position AVEVA not as another closed PLM backbone, but as an open engineering-to-operations decision-control layer. This means linking configuration, change impact, approval evidence, production readiness, handover records and operations feedback above specialist tools, while respecting the customer's existing systems and data ownership. This message is practical for large shipyards because it does not ask them to abandon every specialist tool; it helps them control decisions across the full lifecycle.",
        "This is also where AVEVA's sustainability message becomes concrete. Reduced rework, fewer duplicated documents, clearer configuration baselines, more reliable handover evidence and better lifecycle learning all support more efficient engineering and more responsible use of resources. I believe this is the right way to explain AVEVA's value to marine executives: not as digital transformation language alone, but as measurable improvement in engineering productivity, delivery risk, compliance confidence and lifecycle cost.",
        "4. Contribution to Korea / Japan Marine Growth",
        "After joining, my first priority would be to align quickly with AVEVA's assigned portfolio, Korea/Japan account priorities, Salesforce pipeline discipline, partner and channel structure, regional marketing activities and the collaboration model between Account Managers, Pre-Sales Consultants, Business Development leadership and product teams. I understand that a BDM must create leverage for the sales organisation, not work as an isolated technical specialist.",
        "My contribution would begin with customer discovery. I would help identify and organise marine pain points around design rework, system integration, hull/class configuration, limited digital continuity, compliance workload, production handover, MRO and lifecycle data ownership. These pain points can then be converted into value hypotheses, discovery questions, benchmark assumptions, benefit estimates, Proof of Capability themes and clear follow-up actions.",
        "I would also support executive-level communication by translating technical detail into business language: cost, schedule, risk, readiness, compliance, productivity and sustainability. My experience with ICDs, test plans, VCRM, evidence packages, ECO/BOM changes, WBS schedules, cost estimates, defect logs and acceptance documents gives me the working discipline to make customer discussions specific and credible.",
        "My goal is to help AVEVA grow meaningful Marine Engineering opportunities in Korea and Japan by combining marine-domain credibility, engineering-system understanding, competitive PLM insight and evidence-based customer engagement. I want to contribute to AVEVA's business growth by making its value clear to customers who need practical answers to real shipbuilding lifecycle problems."
    )

    $range = $Doc.Range($heading.Range.End, $Doc.Content.End - 1)
    $range.Text = ($paras -join "`r") + "`r"

    foreach ($p in $Doc.Paragraphs) {
        if ($p.Range.Start -gt $heading.Range.Start) {
            $txt = Normalize-Text $p.Range.Text
            if ($txt.Length -gt 0) {
                $p.Range.Font.Name = "Arial"
                $p.Range.Font.Size = [single]10.7
                $p.Range.Font.Bold = [int]0
                $p.Range.ParagraphFormat.SpaceBefore = [single]0
                $p.Range.ParagraphFormat.SpaceAfter = [single]4
                $p.Range.ParagraphFormat.LineSpacingRule = 0
                if ($txt -match "^[1-4]\. ") {
                    $p.Range.Font.Bold = [int]1
                    $p.Range.Font.Size = [single]11
                    $p.Range.ParagraphFormat.SpaceBefore = [single]8
                    $p.Range.ParagraphFormat.SpaceAfter = [single]5
                }
            }
        }
    }
}

$mainDocx = "D:\이력서\AVEVA - Business Development Manager\Joonha_Lee_AVEVA_BDM_Engineering_Marine_Korea_Japan_Seoul_Resume.docx"
$mainPdf = "D:\이력서\AVEVA - Business Development Manager\Joonha_Lee_AVEVA_BDM_Engineering_Marine_Korea_Japan_Seoul_Resume.pdf"
$docx = "D:\이력서\AVEVA - Business Development Manager\Joonha_Lee_AVEVA_BDM_Engineering_Marine_Korea_Japan_Seoul_Resume_Formal_Tone.docx"
$pdf = "D:\이력서\AVEVA - Business Development Manager\Joonha_Lee_AVEVA_BDM_Engineering_Marine_Korea_Japan_Seoul_Resume_Formal_Tone.pdf"
$tempPdf = "C:\Users\namma\.claude\aveva_bdm_resume_work\self_intro_detailed_export.pdf"

if (Test-Path -LiteralPath $tempPdf) { Remove-Item -LiteralPath $tempPdf -Force }

$word = New-Object -ComObject Word.Application
$word.Visible = $false
$doc = $null
try {
    $doc = $word.Documents.Open($docx, $false, $false)
    Replace-SelfIntroduction -Doc $doc
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
