$ErrorActionPreference = "Stop"

$SourcePath = "D:\이력서\AVEVA - Business Development Manager\Joonha_Lee_AVEVA_BDM_Engineering_Marine_Korea_Japan_Seoul_Resume_Formal_Tone.docx"
$OutPath = "D:\이력서\AVEVA - Business Development Manager\Joonha_Lee_ABB_Portfolio_Industry_Manager_Marine_Ports_Industrial_Automation_Resume.docx"
$PdfPath = "D:\이력서\AVEVA - Business Development Manager\Joonha_Lee_ABB_Portfolio_Industry_Manager_Marine_Ports_Industrial_Automation_Resume.pdf"

Copy-Item -LiteralPath $SourcePath -Destination $OutPath -Force

$replacements = New-Object System.Collections.Generic.List[object]
function Add-Replacement([string]$Old, [string]$New) {
    $replacements.Add([pscustomobject]@{ Old = $Old; New = $New }) | Out-Null
}

Add-Replacement @'
Target: AVEVA - Business Development Manager - Engineering (Marine), Korea / Japan
'@ @'
Target: ABB - Portfolio & Industry Manager / Business Development, Marine & Ports / Industrial Automation
'@

Add-Replacement @'
Marine Engineering & Industrial Software Business Development Professional
'@ @'
Technical Strategy & Portfolio Business Development Professional | Marine & Defense | Industrial Automation
'@

Add-Replacement @'
Marine and industrial engineering professional with over 21 years of experience across ROK Navy shipbuilding programs (FFX Batch-II, Jangbogo-III), defense electronics and a global LEO satellite terminal program. Combines hands-on shipyard lifecycle credibility — requirements/VCRM, E3D/Marine design data, MTO/eBOM, baseline/effectivity control, ECR/ECO and class/handover evidence — with technical value selling that translates marine customer pain points, such as design rework and system integration challenges, into value propositions, benchmark-based benefit cases and qualified opportunities. Ready to support AVEVA regional sales teams, Account Managers and Pre-Sales Consultants across Korea and Japan with marine domain credibility, competitive PLM insight covering Siemens, Dassault, Hexagon, PTC and NAPA, and structured pipeline support aligned with Salesforce CRM.
'@ @'
Senior Technical Strategy & Systems Engineering professional with over 21 years of experience in mission-critical industries, including Marine & Defense Systems, naval shipbuilding, defense electronics and global LEO satellite communications. Adept at translating complex customer requirements and field issue insights into solution portfolio mapping, portfolio positioning, technical sales enablement and qualified business opportunities. Combines first-hand system integration, commissioning, acceptance and lifecycle evidence experience with cross-functional leadership and APAC/global customer interface capability. Ready to help position ABB's Marine & Ports, Industrial Automation & Control, and Power & Energy Infrastructure portfolio for Korea/APAC customers through industry strategy, GTM support and field-to-business value creation.
'@

Add-Replacement @'
Marine Business Development / Pipeline Support
'@ @'
Portfolio Positioning / Industry Strategy & GTM
'@
Add-Replacement @'
Shipbuilding and naval-domain credibility for discovering customer pain points, qualifying Marine Engineering opportunities, supporting lead generation and keeping pipeline actions Salesforce CRM-ready for regional account teams.
'@ @'
Ability to translate mission-critical customer requirements into portfolio positioning, industry strategy, GTM narratives, account hypotheses and CRM-ready pipeline actions for regional sales teams.
'@
Add-Replacement @'
Technical Value Selling / Solution Positioning
'@ @'
Customer Requirement Translation / Solution Portfolio Mapping
'@
Add-Replacement @'
Value-selling language for design rework, configuration inconsistency, limited digital continuity, compliance workload and lifecycle-data gaps.
'@ @'
Technical synthesis capability to map complex customer requirements, field issues and integration constraints into relevant automation, control, energy and lifecycle-service solution themes.
'@
Add-Replacement @'
Marine / Shipbuilding Lifecycle Domain
'@ @'
Marine & Defense Systems / Mission-Critical Infrastructure
'@
Add-Replacement @'
Practical understanding of ROKN programs and the standard large-shipyard lifecycle: requirements/VCRM, basic/detail design, E3D/Marine/Draw, MTO/eBOM, baseline/effectivity, ECR/ECO, class/regulatory evidence and production handover.
'@ @'
Practical understanding of naval shipbuilding, submarine/surface-ship communication systems, C4I-linked interfaces, system integration, FAT/SAT, sea trials, customer acceptance and after-delivery support.
'@
Add-Replacement @'
Customer Engagement / Executive Communication
'@ @'
Technical Sales Enablement / Executive Customer Interface
'@
Add-Replacement @'
Customer meetings, technical reviews, proposal presentations and executive briefings translated into cost, schedule, risk, fleet readiness and lifecycle efficiency.
'@ @'
Customer meetings, technical reviews, proposal presentations and executive briefings translated into cost, schedule, risk, uptime, fleet readiness, lifecycle efficiency and sustainability value.
'@
Add-Replacement @'
Competitive PLM / Engineering Software Insight
'@ @'
Industrial Automation & Control / Power & Energy Infrastructure Adjacency
'@
Add-Replacement @'
Working comparison of AVEVA E3D/Marine, Unified Engineering, AIM and PI/CONNECT against Siemens, Dassault, Hexagon, PTC and NAPA.
'@ @'
Systems-level understanding of control interfaces, industrial Ethernet, fieldbus concepts, shipboard networks, drive/control interaction, power/interface reliability and high-availability infrastructure requirements.
'@
Add-Replacement @'
AVEVA Values / Sustainability / Trust
'@ @'
Lifecycle Business Management / Field-to-Business Value Creation
'@
Add-Replacement @'
Working style aligned with Impact, Aspiration, Curiosity and Trust; sustainability narrative built on reduced rework, digital continuity, evidence-based decisions and responsible use of engineering resources.
'@ @'
Working style focused on evidence-based trust, lifecycle issue closure, customer risk reduction, service opportunity identification, portfolio feedback and responsible use of engineering and energy resources.
'@

Add-Replacement @'
Led customer-facing SIT, verification governance and release-readiness control for a global LEO terminal program; managed 400+ verification items.
'@ @'
Led customer-facing SIT, verification governance and release-readiness control for a global LEO terminal program; converted 400+ verification items into customer approval evidence and lifecycle risk insight.
'@
Add-Replacement @'
Managed defense/aerospace RFPs, tender responses, cost/schedule planning, technical reviews and customer/partner coordination.
'@ @'
Managed defense/aerospace RFPs, tender responses, cost/schedule planning, technical reviews and customer/partner coordination; translated requirements into proposal strategy and executable delivery plans.
'@
Add-Replacement @'
Built 15 years of naval/shipbuilding delivery credibility across requirements, design, FAT/SAT, sea trial and handover; mapped shipyard pain points to standard large-shipyard lifecycle controls.
'@ @'
Built 15 years of naval/shipbuilding delivery credibility across requirements, design, installation, FAT/SAT, sea trial and handover; mapped shipyard pain points to system integration, automation/control and lifecycle-service value themes.
'@

Add-Replacement @'
Job-Fit Map: AVEVA Marine BDM Requirements vs. Career Evidence
'@ @'
Job-Fit Map: ABB Portfolio / Industry Manager Requirements vs. Career Evidence
'@
Add-Replacement @'
AVEVA BDM Requirement
'@ @'
ABB Role Requirement
'@
Add-Replacement @'
Value to AVEVA / Client
'@ @'
Value to ABB / Client
'@
Add-Replacement @'
Grow Marine pipeline and qualify opportunities
'@ @'
Portfolio positioning and industry strategy / GTM
'@
Add-Replacement @'
Supported RFP analysis, technical proposals, customer meetings and issue prioritisation across naval, defence and satellite programs.
'@ @'
Supported RFP analysis, technical proposals, customer meetings and issue prioritisation across naval, defence and satellite programs; built value narratives from real customer constraints.
'@
Add-Replacement @'
Converts Korea/Japan marine pain points into qualified Engineering opportunities and CRM-ready next actions.
'@ @'
Converts Korea/APAC marine, defense and infrastructure pain points into industry strategy, solution themes, qualified opportunities and CRM-ready next actions.
'@
Add-Replacement @'
Bring marine domain and competitive insight
'@ @'
Translate customer requirements into solution portfolio mapping
'@
Add-Replacement @'
15 years in ROK Navy and shipyard programs; practical view of requirements/VCRM, E3D/Marine/Draw, MTO/eBOM, baseline/effectivity, ECR/ECO and handover evidence; benchmarked AVEVA against Siemens, Dassault, Hexagon, PTC and NAPA.
'@ @'
15 years in ROK Navy and shipyard programs plus global satellite SIT; practical view of requirements, interfaces, FAT/SAT, commissioning, acceptance evidence, ECO/BOM impact and field issue closure.
'@
Add-Replacement @'
Creates credible discovery with shipyard, offshore and naval customers, grounded in real lifecycle process.
'@ @'
Creates credible discovery with shipyard, marine, defense, energy and infrastructure customers, grounded in real system integration and lifecycle process.
'@
Add-Replacement @'
Build value propositions and benefit cases
'@ @'
Build technical sales enablement and benefit cases
'@
Add-Replacement @'
Managed 400+ verification items, VCRM-style evidence, ECO/BOM impact, WBS schedules, cost estimates and acceptance criteria.
'@ @'
Managed 400+ verification items, evidence packages, ECO/BOM impact, WBS schedules, cost estimates, acceptance criteria and customer-facing issue justification.
'@
Add-Replacement @'
Supports benefit cases for lower rework, faster change-impact review, stronger configuration consistency and compliance efficiency.
'@ @'
Supports benefit cases for lower integration risk, improved uptime/reliability, faster issue closure, stronger lifecycle service readiness and more confident investment decisions.
'@
Add-Replacement @'
Engage executives in business language
'@ @'
Engage executives and cross-functional stakeholders in business language
'@
Add-Replacement @'
Translates technical detail into cost, schedule, risk, readiness and sustainability value.
'@ @'
Translates technical detail into cost, schedule, risk, uptime, readiness, energy efficiency and sustainability value.
'@
Add-Replacement @'
Coordinate sales, pre-sales, partners and delivery teams
'@ @'
Lead cross-functional customer, sales, product and delivery interfaces
'@
Add-Replacement @'
Fits regional account teams that require disciplined follow-up across sales, pre-sales, partners and customer functions.
'@ @'
Fits regional teams that require disciplined follow-up across sales, portfolio/product, partners, service, engineering and customer functions.
'@
Add-Replacement @'
Position AVEVA Marine PLM / Engineering differentiation
'@ @'
Convert field issues into lifecycle business and portfolio feedback
'@
Add-Replacement @'
Defined an AVEVA narrative around an open decision layer linking configuration, impact, evidence, approval and operations feedback.
'@ @'
Built a repeatable way to classify field issues by customer impact, root cause, corrective action, re-verification evidence and lifecycle follow-up opportunity.
'@
Add-Replacement @'
Positions AVEVA above document-centric workflows as the engineering-to-operations decision loop for marine customers.
'@ @'
Helps ABB convert installed-base issues and site learnings into lifecycle service value, portfolio optimization and trusted long-term customer engagement.
'@

Add-Replacement @'
Established a project-control discipline linking customer requirements, field issues, defect history, test evidence and approval criteria. The same discipline applies to Marine BDM work: turning customer pain points into solution fit, benefit evidence and qualified next actions.
'@ @'
Established a project-control discipline linking customer requirements, field issues, defect history, test evidence and approval criteria. The same discipline applies to ABB portfolio and industry management work: turning customer pain points into solution fit, technical sales evidence and qualified next actions.
'@
Add-Replacement @'
Built practical experience in requirements-to-evidence traceability, which is closely aligned with PLM-based lifecycle data control and customer acceptance management.
'@ @'
Built practical experience in requirements-to-evidence traceability, which is closely aligned with lifecycle business management, customer acceptance and mission-critical infrastructure governance.
'@
Add-Replacement @'
Built a practical working style suitable for customer technical support, PLM issue tracking, and product development feedback loops.
'@ @'
Built a practical working style suitable for customer technical support, lifecycle issue tracking, product-development feedback and field-to-business opportunity conversion.
'@
Add-Replacement @'
Gained relevant experience for AVEVA-style customer issue resolution between customers, product development, and portfolio or strategy teams.
'@ @'
Gained relevant experience for ABB-style customer issue resolution between customers, product development, service, portfolio and strategy teams.
'@

Add-Replacement @'
Managed marine system integration, technical documentation, FAT/SAT, sea trial support, customer acceptance, field issue response and coordination with Navy, shipyards, inspectors, subcontractors, production, quality and engineering teams. Mapped this delivery experience to large-shipyard lifecycle controls: requirements/VCRM, basic/detail design, E3D/Marine/Draw, MTO/eBOM, hull/block baseline, ECR/ECO, production release and handover evidence.
'@ @'
Managed marine system integration, technical documentation, FAT/SAT, sea trial support, customer acceptance, field issue response and coordination with Navy, shipyards, inspectors, subcontractors, production, quality and engineering teams. Mapped this delivery experience to ABB-relevant value themes: mission-critical uptime, automation/control interface reliability, power and energy infrastructure resilience, lifecycle service readiness and trusted customer acceptance.
'@
Add-Replacement @'
Built experience directly relevant to AVEVA Marine PLM consulting, where customer needs must be translated into solution architecture, implementation scope, and technical proposal materials.
'@ @'
Built experience directly relevant to ABB portfolio and industry strategy, where customer needs must be translated into solution portfolio mapping, implementation scope, value narrative and technical proposal materials.
'@
Add-Replacement @'
Developed system architecture and interface-management experience relevant to next-generation shipbuilding platforms and marine PLM environments.
'@ @'
Developed system architecture and interface-management experience relevant to next-generation shipbuilding platforms, industrial automation interfaces and mission-critical marine infrastructure.
'@
Add-Replacement @'
Built a strong foundation in lifecycle documentation, interface traceability, and evidence-based acceptance control, which is highly relevant to Marine PLM implementation and consulting.
'@ @'
Built a strong foundation in lifecycle documentation, interface traceability and evidence-based acceptance control, which is highly relevant to automation/control portfolio positioning and lifecycle business management.
'@

Add-Replacement @'
AVEVA Marine PLM / Engineering Data / Digital Thread working knowledge (self-built competitive strategy deck and KPI model).
'@ @'
ABB Marine & Ports / Industrial Automation portfolio-positioning working knowledge; self-built strategy narrative and KPI model for mission-critical marine and infrastructure customers.
'@
Add-Replacement @'
ready to support Korea/Japan regional opportunities in English.
'@ @'
ready to support Korea/APAC regional opportunities in English.
'@
Add-Replacement @'
AVEVA E3D/Marine concepts, Unified Engineering, AIM, PI/CONNECT positioning, PLM, Digital Thread, Digital Twin, Smart Shipyard, ERP/MES/PLM boundaries, EPLAN concept, HART/Modbus/Fieldbus/Industrial Ethernet, Salesforce CRM awareness, ICD, VCRM, ECO/BOM, WBS, cost estimate, defect log, evidence package, Wireshark, iperf, Linux, VMware and RF measurement equipment.
'@ @'
ABB Marine & Ports portfolio awareness, Industrial Automation & Control, Power & Energy Infrastructure, drive/control interface awareness, PLC/control-system concepts, HART/Modbus/Fieldbus/Industrial Ethernet, System Integration, Mission-Critical Infrastructure, Digital Thread, Digital Twin, Smart Shipyard, lifecycle business management, Salesforce CRM awareness, ICD, VCRM, ECO/BOM, WBS, cost estimate, defect log, evidence package, Wireshark, iperf, Linux, VMware and RF measurement equipment.
'@

Add-Replacement @'
1. Motivation and AVEVA Fit
'@ @'
1. Motivation and ABB Fit
'@
Add-Replacement @'
I am applying for the Business Development Manager - Engineering (Marine), Korea / Japan role in Seoul because it requires the exact combination I have built throughout my career: marine engineering credibility, customer-facing technical judgement, and the ability to translate complex project issues into business value. This role is not only about selling software; it is about understanding how shipyards, owners, engineering teams and executives make decisions when schedule, quality, compliance and lifecycle risk are all connected.
'@ @'
I am applying for ABB's Portfolio / Industry Management and Business Development opportunity because it requires the exact combination I have built throughout my career: mission-critical engineering credibility, customer-facing technical judgement, and the ability to translate complex technical requirements into portfolio strategy, solution mapping and business value. This role is not simply about product explanation; it is about understanding how marine, defense, energy and infrastructure customers make investment decisions when uptime, safety, compliance, lifecycle cost and operational risk are all connected.
'@
Add-Replacement @'
That experience is the reason AVEVA is highly attractive to me. Marine customers do not purchase software functions in isolation. They are looking for lower design rework, stronger configuration control, faster change-impact review, reliable approval evidence, better production handover and a trusted digital thread from engineering through build, operation and lifecycle learning. AVEVA Engineering / Marine is positioned to turn those engineering problems into measurable business outcomes.
'@ @'
That experience is the reason ABB is highly attractive to me. Marine and infrastructure customers do not purchase automation, control, electrification or lifecycle services in isolation. They are looking for safer operations, higher uptime, stronger system integration, faster issue recovery, lower lifecycle risk, improved energy efficiency and trusted local support. ABB's Marine & Ports, Industrial Automation & Control, and Power & Energy Infrastructure portfolio is positioned to turn these operational problems into measurable business outcomes.
'@
Add-Replacement @'
I understand AVEVA values Impact, Aspiration, Curiosity and Trust. These values are close to my working style. In complex projects, I try to make practical impact, understand the customer goal, ask why before deciding, and build trust through clear evidence and follow-up. I also respect AVEVA's sustainability direction: connected engineering data should reduce waste, improve productivity, support better decisions and help customers use industrial resources more responsibly.
'@ @'
I understand ABB's emphasis on sustainable, reliable and efficient industrial progress. This direction is close to my working style. In complex projects, I try to make practical impact, understand the customer's operational goal, ask why before deciding, and build trust through clear evidence and follow-up. I also respect ABB's sustainability message: automation, electrification and lifecycle services should reduce waste, improve productivity, support safer decisions and help customers use energy and industrial resources more responsibly.
'@
Add-Replacement @'
2. Marine Customer and Lifecycle Expertise
'@ @'
2. Customer Requirement Translation and Mission-Critical Domain Expertise
'@
Add-Replacement @'
My core strength is the ability to stand between the customer, engineering teams and business stakeholders, then convert technical reality into a clear action plan. In shipbuilding and defence projects, I learned that the real issue is often not a single defect or requirement. It is the lack of a shared view across configuration, interface, schedule, test evidence, approval status and downstream impact. This is exactly where a strong Marine BDM must create value.
'@ @'
My core strength is the ability to stand between the customer, engineering teams and business stakeholders, then convert technical reality into a clear action plan. In shipbuilding and defense projects, I learned that the real issue is often not a single defect or requirement. It is the lack of a shared view across interfaces, operating conditions, schedule risk, test evidence, acceptance status and downstream lifecycle impact. This is exactly where a strong technical strategy and portfolio manager must create value.
'@
Add-Replacement @'
I can connect that experience to the standard large-shipyard lifecycle in Korea and Japan (for example HD Hyundai-class yards): requirements/VCRM, basic and detail design, E3D/Marine/Draw, MTO/eBOM, baseline and effectivity control, ECR/ECO, production release, classification approval evidence and handover. This process view is important for AVEVA because many customer conversations begin with symptoms such as document inconsistency, delayed change review, unclear ownership of engineering data, weak handover evidence or fragmented digital continuity.
'@ @'
I can connect that experience to ABB-relevant customer environments in Korea and APAC: shipyards, naval programs, ports, onboard systems, communication/control infrastructure, power and energy interfaces, and high-availability industrial operations. This process view is important because many customer conversations begin with symptoms such as recurring field issues, unclear root cause ownership, delayed acceptance, interface instability, lifecycle service gaps, energy-efficiency pressure or fragmented operational data.
'@
Add-Replacement @'
At Intellian Technologies, I expanded this discipline in a global customer environment by managing requirements, field issues, defect history, ECO/BOM changes, verification evidence and release judgement for a LEO satellite-terminal program. More than 400 verification checklist items had to be controlled across RF/antenna, control electronics, network interface, environmental verification and production verification. That experience reinforced a principle I would bring to AVEVA customer engagement: trust is built through evidence, not claims.
'@ @'
At Intellian Technologies, I expanded this discipline in a global customer environment by managing requirements, field issues, defect history, ECO/BOM changes, verification evidence and release judgement for a LEO satellite-terminal program. More than 400 verification checklist items had to be controlled across RF/antenna, control electronics, network interface, environmental verification and production verification. That experience reinforced a principle I would bring to ABB customer engagement: trust is built through evidence, not claims.
'@
Add-Replacement @'
Commercially, this background fits the technical value-selling side of the BDM role. I can support Account Managers and Pre-Sales Consultants by structuring discovery questions, clarifying customer pain points, preparing value-proposition inputs, shaping Proof of Capability storylines and turning technical discussion into CRM-ready next actions. My contribution is strongest where customer needs must be interpreted accurately before a solution can be positioned credibly.
'@ @'
Commercially, this background fits the technical value-selling and portfolio-positioning side of the role. I can support Account Managers, Product/Portfolio teams and Pre-Sales or Service specialists by structuring discovery questions, clarifying customer pain points, preparing value-proposition inputs, shaping proof-of-value storylines and turning technical discussion into CRM-ready next actions. My contribution is strongest where customer needs must be interpreted accurately before a solution can be positioned credibly.
'@
Add-Replacement @'
3. AVEVA Marine PLM Perspective
'@ @'
3. Portfolio / Industry Strategy Perspective
'@
Add-Replacement @'
I have studied AVEVA Marine / PLM in depth - from the viewpoints of product position, competitor strengths, customer pain points and future strategy - and consolidated that work into my own executive strategy deck, including a KPI-gated delivery and acceptance model covering configuration accuracy, change-impact recall, evidence completeness and approval response time. That work strengthened my motivation to join AVEVA because I could see a clear strategic opportunity in the marine industry, and it gives me concrete material for benchmark and benefit-estimation discussions with customers.
'@ @'
I have studied marine and industrial portfolio strategy from the viewpoints of product position, competitor strengths, customer pain points and future go-to-market logic, and consolidated that work into my own executive strategy materials and KPI-based value model. That work strengthened my motivation to move from pure engineering delivery into technical strategy and business development, because I can see clear strategic opportunities in marine, ports, industrial automation and power/energy infrastructure customers.
'@
Add-Replacement @'
My practical comparison is as follows. Siemens Teamcenter is strong in configuration, BOM and change governance. Dassault 3DEXPERIENCE is strong in collaboration and virtual-twin positioning. Hexagon has strength in asset lifecycle and operations context. PTC is strong in requirements, ALM and traceability patterns. NAPA is trusted in naval architecture and stability calculation. AVEVA's opportunity is different: it can connect engineering truth, lifecycle context and operational feedback through E3D/Marine, Unified Engineering, AIM, PI/CONNECT and open integration, strengthened by AVEVA's public AI direction such as the CONNECT Industrial AI Assistant and Generative / Predictive Design AI. The white space is clear: no current vendor owns the complete loop from design change to naval calculation, evidence, approval and operations feedback, and AVEVA is the best-positioned company to own it.
'@ @'
My practical perspective is that ABB can win when its portfolio is positioned around customer outcomes rather than isolated hardware or software functions. Marine and infrastructure customers care about mission continuity, automation/control reliability, energy efficiency, maintainability, cybersecurity-aware connectivity, lifecycle service response and local execution credibility. My background allows me to connect ABB's portfolio language to real customer conditions: shipboard interfaces, commissioning constraints, acceptance evidence, field issue closure, after-delivery support and executive-level business risk.
'@
Add-Replacement @'
For Korea and Japan marine customers, I would position AVEVA not as another closed PLM backbone, but as an open engineering-to-operations decision-control layer. This means linking configuration, change impact, approval evidence, production readiness, handover records and operations feedback above specialist tools, while respecting the customer's existing systems and data ownership. This message is practical for large shipyards because it does not ask them to abandon every specialist tool; it helps them control decisions across the full lifecycle.
'@ @'
For Korea/APAC customers, I would position ABB as a trusted technology partner that connects reliable automation, control, electrification, digital lifecycle insight and field-service credibility. This means mapping customer requirements to the right portfolio themes, explaining integration and lifecycle value in business language, and feeding customer field insights back into portfolio optimization and GTM priorities.
'@
Add-Replacement @'
This is also where AVEVA's sustainability message becomes concrete. Reduced rework, fewer duplicated documents, clearer configuration baselines, more reliable handover evidence and better lifecycle learning all support more efficient engineering and more responsible use of resources. I believe this is the right way to explain AVEVA's value to marine executives: not as digital transformation language alone, but as measurable improvement in engineering productivity, delivery risk, compliance confidence and lifecycle cost.
'@ @'
This is also where ABB's sustainability message becomes concrete. Higher uptime, efficient energy use, fewer repeated failures, better lifecycle service planning and stronger evidence-based decisions all support safer operations and more responsible use of industrial resources. I believe this is the right way to explain ABB's value to executives: not as technology language alone, but as measurable improvement in operational reliability, delivery risk, compliance confidence, energy efficiency and lifecycle cost.
'@
Add-Replacement @'
4. Contribution to Korea / Japan Marine Growth
'@ @'
4. Contribution to Korea / APAC Marine, Automation and Infrastructure Growth
'@
Add-Replacement @'
After joining, my first priority would be to align quickly with AVEVA's assigned portfolio, Korea/Japan account priorities, Salesforce pipeline discipline, partner and channel structure, regional marketing activities and the collaboration model between Account Managers, Pre-Sales Consultants, Business Development leadership and product teams. I understand that a BDM must create leverage for the sales organisation, not work as an isolated technical specialist.
'@ @'
After joining, my first priority would be to align quickly with ABB's assigned portfolio, Korea/APAC account priorities, pipeline discipline, partner and channel structure, regional marketing activities and the collaboration model between Account Managers, Portfolio/Product teams, Service teams, Business Development leadership and technical specialists. I understand that a portfolio or industry manager must create leverage for the sales organization, not work as an isolated technical specialist.
'@
Add-Replacement @'
My contribution would begin with customer discovery. I would help identify and organise marine pain points around design rework, system integration, hull/class configuration, limited digital continuity, compliance workload, production handover, MRO and lifecycle data ownership. These pain points can then be converted into value hypotheses, discovery questions, benchmark assumptions, benefit estimates, Proof of Capability themes and clear follow-up actions.
'@ @'
My contribution would begin with customer discovery. I would help identify and organize marine, port, defense, energy and industrial pain points around system integration, automation/control reliability, power-interface resilience, field issue recurrence, commissioning risk, lifecycle service readiness, energy efficiency and operational data ownership. These pain points can then be converted into value hypotheses, discovery questions, benchmark assumptions, benefit estimates, proof-of-value themes and clear follow-up actions.
'@
Add-Replacement @'
I would also support executive-level communication by translating technical detail into business language: cost, schedule, risk, readiness, compliance, productivity and sustainability. My experience with ICDs, test plans, VCRM, evidence packages, ECO/BOM changes, WBS schedules, cost estimates, defect logs and acceptance documents gives me the working discipline to make customer discussions specific and credible.
'@ @'
I would also support executive-level communication by translating technical detail into business language: cost, schedule, risk, uptime, readiness, compliance, productivity, energy efficiency and sustainability. My experience with ICDs, test plans, VCRM-style evidence packages, ECO/BOM changes, WBS schedules, cost estimates, defect logs and acceptance documents gives me the working discipline to make customer discussions specific and credible.
'@
Add-Replacement @'
My goal is to help AVEVA grow meaningful Marine Engineering opportunities in Korea and Japan by combining marine-domain credibility, engineering-system understanding, competitive PLM insight and evidence-based customer engagement, while making AVEVA's value clear to customers who need practical answers to real shipbuilding lifecycle problems.
'@ @'
My goal is to help ABB grow meaningful Marine & Ports, Industrial Automation and Power/Energy Infrastructure opportunities in Korea/APAC by combining mission-critical domain credibility, system-integration understanding, portfolio-positioning discipline and evidence-based customer engagement, while making ABB's value clear to customers who need practical answers to real operational and lifecycle problems.
'@

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
        $app = New-Object -ComObject Word.Application
        return $app
    }
    catch {
        Start-Process -FilePath "C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE" -ArgumentList "/automation" -WindowStyle Hidden
        Start-Sleep -Seconds 5
        $app = New-Object -ComObject Word.Application
        return $app
    }
}

$word = New-WordApplication
$word.Visible = $false
$doc = $null

try {
    $doc = $word.Documents.Open($OutPath, $false, $false)

    $hits = 0
    $misses = New-Object System.Collections.Generic.List[string]
    foreach ($rep in $replacements) {
        $target = Normalize-Text $rep.Old
        $found = $false
        foreach ($p in $doc.Paragraphs) {
            $current = Normalize-Text $p.Range.Text
            if ($current -eq $target) {
                Set-RangeTextKeepMark $p.Range $rep.New
                $hits++
                $found = $true
                break
            }
        }
        if (-not $found) {
            $misses.Add($target) | Out-Null
        }
    }

    $doc.Save()
    $doc.ExportAsFixedFormat($PdfPath, 17)
    $wordCount = $doc.ComputeStatistics(0)
    $pageCount = $doc.ComputeStatistics(2)
    $doc.Close([ref]$true)
    $doc = $null

    Write-Host "Saved DOCX: $OutPath"
    Write-Host "Saved PDF : $PdfPath"
    Write-Host "Pages     : $pageCount"
    Write-Host "Words     : $wordCount"
    Write-Host "Hits      : $hits"
    Write-Host "Misses    : $($misses.Count)"
    foreach ($m in $misses) {
        Write-Host "MISS: $m"
    }
}
finally {
    if ($doc -ne $null) {
        try { $doc.Close([ref]$false) } catch {}
    }
    if ($word -ne $null -and $word.Documents.Count -eq 0) {
        $word.Quit()
    }
}
