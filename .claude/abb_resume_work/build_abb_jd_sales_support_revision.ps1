$ErrorActionPreference = "Stop"

$SourcePath = "D:\이력서\ABB - Portfolio and Industry Manager\Joonha_Lee_ABB_Portfolio_Industry_Manager_Marine_Ports_Industrial_Automation_Resume.docx"
$OutPath = "D:\이력서\ABB - Portfolio and Industry Manager\Joonha_Lee_ABB_Portfolio_Industry_Manager_Marine_Ports_Industrial_Automation_Resume_JD_SalesSupport_Revised.docx"
$PdfPath = "D:\이력서\ABB - Portfolio and Industry Manager\Joonha_Lee_ABB_Portfolio_Industry_Manager_Marine_Ports_Industrial_Automation_Resume_JD_SalesSupport_Revised.pdf"

Copy-Item -LiteralPath $SourcePath -Destination $OutPath -Force

$replacements = New-Object System.Collections.Generic.List[object]
function Add-Replacement([string]$Old, [string]$New) {
    $replacements.Add([pscustomobject]@{ Old = $Old; New = $New }) | Out-Null
}

Add-Replacement @'
Target: ABB - Portfolio & Industry Manager / Business Development, Marine & Ports / Industrial Automation
'@ @'
Target: ABB - Portfolio & Industry Manager, Machine Automation / Industrial Automation, South Korea
'@

Add-Replacement @'
Technical Strategy & Portfolio Business Development Professional | Marine & Defense | Industrial Automation
'@ @'
Technical Strategy & Portfolio Business Development Professional | Machine Automation | Marine & Defense | Industrial Automation
'@

Add-Replacement @'
Senior Technical Strategy & Systems Engineering professional with over 21 years of experience in mission-critical industries, including Marine & Defense Systems, naval shipbuilding, defense electronics and global LEO satellite communications. Adept at translating complex customer requirements and field issue insights into solution portfolio mapping, portfolio positioning, technical sales enablement and qualified business opportunities. Combines first-hand system integration, commissioning, acceptance and lifecycle evidence experience with cross-functional leadership and APAC/global customer interface capability. Ready to help position ABB's Marine & Ports, Industrial Automation & Control, and Power & Energy Infrastructure portfolio for Korea/APAC customers through industry strategy, GTM support and field-to-business value creation.
'@ @'
Senior Technical Strategy & Systems Engineering professional with over 21 years of experience in mission-critical industries, including Marine & Defense Systems, naval shipbuilding, defense electronics and global LEO satellite communications. Adept at structuring complex customer problems, eliciting unspoken technical requirements, and translating field issues into offering strategy, solution portfolio mapping, go-to-market narratives, sales enablement tools and qualified business opportunities. Combines first-hand system integration, commissioning, acceptance, cost-estimation and lifecycle evidence experience with cross-functional leadership across development, quality, production, sales and customer interfaces. Ready to support ABB Machine Automation and Industrial Automation growth in Korea by aligning portfolio strategy, segment strategy, market intelligence, roadmap planning and lifecycle management with measurable customer value.
'@

Add-Replacement @'
Ability to translate mission-critical customer requirements into portfolio positioning, industry strategy, GTM narratives, account hypotheses and CRM-ready pipeline actions for regional sales teams.
'@ @'
Ability to translate mission-critical customer requirements into offering strategy, portfolio positioning, segment strategy, GTM narratives, account hypotheses and CRM-ready pipeline actions for regional sales teams.
'@

Add-Replacement @'
Technical synthesis capability to map complex customer requirements, field issues and integration constraints into relevant automation, control, energy and lifecycle-service solution themes.
'@ @'
Technical synthesis capability to map complex customer requirements, unspoken needs, field issues and integration constraints into relevant automation, control, energy and lifecycle-service solution themes.
'@

Add-Replacement @'
Customer meetings, technical reviews, proposal presentations and executive briefings translated into cost, schedule, risk, uptime, fleet readiness, lifecycle efficiency and sustainability value.
'@ @'
Customer meetings, technical reviews, proposal presentations, cost-estimation support and executive briefings translated into cost, schedule, risk, uptime, product profitability, lifecycle efficiency and sustainability value.
'@

Add-Replacement @'
Working style focused on evidence-based trust, lifecycle issue closure, customer risk reduction, service opportunity identification, portfolio feedback and responsible use of engineering and energy resources.
'@ @'
Working style focused on evidence-based trust, root-cause analysis, recurrence-prevention structures, lifecycle issue closure, service opportunity identification, portfolio feedback and responsible use of engineering and energy resources.
'@

Add-Replacement @'
Built 15 years of naval/shipbuilding delivery credibility across requirements, design, installation, FAT/SAT, sea trial and handover; mapped shipyard pain points to system integration, automation/control and lifecycle-service value themes.
'@ @'
Built 15 years of naval/shipbuilding delivery credibility across requirements, design, installation, FAT/SAT, sea trial and handover; supported Sales HQ on every new special-vessel opportunity with cost estimation, proposal inputs and value narratives, including a KRW 1.2B Indonesia submarine project win.
'@

Add-Replacement @'
Portfolio positioning and industry strategy / GTM
'@ @'
Offering strategy, portfolio positioning and industry strategy / GTM
'@

Add-Replacement @'
Supported RFP analysis, technical proposals, customer meetings and issue prioritisation across naval, defence and satellite programs; built value narratives from real customer constraints.
'@ @'
Supported RFP analysis, technical proposals, customer meetings, cost estimates and issue prioritisation across naval, defence and satellite programs; built value narratives from real customer constraints and unspoken technical requirements.
'@

Add-Replacement @'
Converts Korea/APAC marine, defense and infrastructure pain points into industry strategy, solution themes, qualified opportunities and CRM-ready next actions.
'@ @'
Converts Korea/APAC industrial, marine, defense and infrastructure pain points into segment strategy, offering themes, qualified opportunities and CRM-ready next actions.
'@

Add-Replacement @'
Managed 400+ verification items, evidence packages, ECO/BOM impact, WBS schedules, cost estimates, acceptance criteria and customer-facing issue justification.
'@ @'
Managed 400+ verification items, evidence packages, ECO/BOM impact, WBS schedules, cost estimates, acceptance criteria, customer-facing issue justification and recurrence-prevention actions.
'@

Add-Replacement @'
Coordinated engineering, quality, production, subcontractors, agencies and global partners through long-cycle delivery and acceptance work.
'@ @'
Coordinated engineering, quality, production, sales, subcontractors, agencies and global partners through long-cycle delivery, sales support and acceptance work.
'@

Add-Replacement @'
Fits regional teams that require disciplined follow-up across sales, portfolio/product, partners, service, engineering and customer functions.
'@ @'
Fits regional teams that require disciplined follow-up across sales, portfolio/product, partners, service, engineering, production and customer functions.
'@

Add-Replacement @'
Delivered ROK Navy ship and submarine communication systems, shipboard broadcasting systems, internal networks, CCTV/VOD and C4I-linked systems from requirements analysis through design, installation, commissioning, acceptance and after-delivery support. Developed first-hand understanding of Korean shipyard workflows, naval customer language, system integration constraints and lifecycle documentation.
'@ @'
Delivered ROK Navy ship and submarine communication systems, shipboard broadcasting systems, internal networks, CCTV/VOD and C4I-linked systems from requirements analysis through design, installation, commissioning, acceptance and after-delivery support. Developed first-hand understanding of Korean shipyard workflows, naval customer language, system integration constraints, lifecycle documentation and new special-vessel sales support.
'@

Add-Replacement @'
Managed marine system integration, technical documentation, FAT/SAT, sea trial support, customer acceptance, field issue response and coordination with Navy, shipyards, inspectors, subcontractors, production, quality and engineering teams. Mapped this delivery experience to ABB-relevant value themes: mission-critical uptime, automation/control interface reliability, power and energy infrastructure resilience, lifecycle service readiness and trusted customer acceptance.
'@ @'
Managed marine system integration, technical documentation, FAT/SAT, sea trial support, customer acceptance, field issue response and coordination with Navy, shipyards, inspectors, subcontractors, production, quality, engineering and sales teams. For every new special-vessel sales opportunity during Daeyang Electric tenure, supported Sales HQ by translating customer requirements into technical configuration, scope, development cost, manpower, material/outsourcing cost, quotation logic and proposal content. Mapped this delivery and sales-support experience to ABB-relevant value themes: offering strategy, time-to-market, sales enablement, product profitability, lifecycle service readiness and trusted customer acceptance.
'@

Add-Replacement @'
(2) Requirements Analysis, Customer Pain-Point Discovery and Proposal Planning
'@ @'
(2) Requirements Analysis, Latent Need Discovery, Cost Estimation and Proposal Planning
'@

Add-Replacement @'
Analysed customer requirements from the Navy, shipyards, and defence authorities and translated them into system architecture, network configuration, equipment specifications, ICDs, technical proposals, and execution plans.
'@ @'
Analysed customer requirements from the Navy, shipyards, and defence authorities, including requirements that customers could not fully articulate, and translated them into system architecture, network configuration, equipment specifications, ICDs, technical proposals, cost-estimation inputs and execution plans.
'@

Add-Replacement @'
Supported proposal presentations and planning documents with a strong focus on operational concept, onboard installation constraints, and real shipboard conditions.
'@ @'
Supported proposal presentations and planning documents with a strong focus on operational concept, onboard installation constraints, real shipboard conditions, development scope, expected manpower, material/outsourcing cost and delivery risk.
'@

Add-Replacement @'
Converted customer needs into implementable system configurations and milestone-based execution plans.
'@ @'
Converted customer needs into implementable system configurations, quotation assumptions, milestone-based execution plans and sales-ready value propositions.
'@

Add-Replacement @'
Built experience directly relevant to ABB portfolio and industry strategy, where customer needs must be translated into solution portfolio mapping, implementation scope, value narrative and technical proposal materials.
'@ @'
Built experience directly relevant to ABB portfolio and industry strategy, where customer needs must be translated into offering strategy, solution portfolio mapping, implementation scope, value narrative, sales tools and technical proposal materials.
'@

Add-Replacement @'
Managed WBS-based schedule plans, budget and manpower allocation, equipment deployment, subcontractor coordination, risk registers, and change control systems.
'@ @'
Managed WBS-based schedule plans, budget and manpower allocation, development cost assumptions, equipment deployment, subcontractor coordination, risk registers, and change control systems.
'@

Add-Replacement @'
Identified schedule risks early and aligned purchasing, production, quality, and subcontractor actions to keep projects within delivery and acceptance targets.
'@ @'
Identified schedule and cost risks early and aligned sales, purchasing, production, quality, engineering and subcontractor actions to keep projects within quotation, delivery and acceptance targets.
'@

Add-Replacement @'
Managed multiple commitments and competing priorities across engineering, production, shipyard, customer, and inspection stakeholders.
'@ @'
Managed multiple commitments and competing priorities across engineering, production, sales, shipyard, customer, and inspection stakeholders, acting as a translator between technical reality and business commitments.
'@

Add-Replacement @'
Led and supported FAT, SAT, commissioning, sea trial, and customer acceptance activities with the Navy, shipbuilders, inspectors, and subcontractors.
'@ @'
Led and supported FAT, SAT, commissioning, sea trial, and customer acceptance activities with the Navy, shipbuilders, inspectors, and subcontractors, feeding site learnings back into recurrence-prevention structures and future quotation assumptions.
'@

Add-Replacement @'
Minimized project closing risk by resolving issues through cause analysis and corrective action rather than responsibility disputes.
'@ @'
Minimized project closing risk by resolving issues through root-cause analysis, corrective action and recurrence-prevention structures rather than responsibility disputes.
'@

Add-Replacement @'
Supported overseas submarine field technical assistance in Indonesia and gained practical experience in customer-facing troubleshooting under site constraints.
'@ @'
Supported overseas submarine field technical assistance in Indonesia and gained practical experience in customer-facing troubleshooting, local-site requirement discovery and sales-to-delivery follow-through under site constraints.
'@

Add-Replacement @'
Worked directly with overseas customers, shipyards, and partners during field issue response and acceptance support.
'@ @'
Worked directly with overseas customers, shipyards, partners and internal sales stakeholders during field issue response, acceptance support and business follow-up.
'@

Add-Replacement @'
Developed the flexibility and professional attitude required for customer visits, on-site support, and potential international travel.
'@ @'
Contributed to an approximately KRW 1.2B Indonesia submarine entertainment/broadcasting system opportunity by proposing a wireless network architecture, supporting cost estimation, persuading customer/shipyard stakeholders, and linking technology to business opportunity.
'@

Add-Replacement @'
ABB Marine & Ports portfolio awareness, Industrial Automation & Control, Power & Energy Infrastructure, drive/control interface awareness, PLC/control-system concepts, HART/Modbus/Fieldbus/Industrial Ethernet, System Integration, Mission-Critical Infrastructure, Digital Thread, Digital Twin, Smart Shipyard, lifecycle business management, Salesforce CRM awareness, ICD, VCRM, ECO/BOM, WBS, cost estimate, defect log, evidence package, Wireshark, iperf, Linux, VMware and RF measurement equipment.
'@ @'
ABB Machine Automation / Industrial Automation portfolio awareness, Marine & Ports portfolio awareness, offering strategy, product-line portfolio strategy, go-to-market model, segment strategy, market intelligence, roadmap planning, lifecycle management, sales enablement, value proposition, competitive benchmarking, drive/control interface awareness, PLC/control-system concepts, HART/Modbus/Fieldbus/Industrial Ethernet, System Integration, Mission-Critical Infrastructure, Digital Thread, Digital Twin, Smart Shipyard, Salesforce CRM awareness, ICD, VCRM, ECO/BOM, WBS, cost estimate, quotation logic, defect log, evidence package, Wireshark, iperf, Linux, VMware and RF measurement equipment.
'@

Add-Replacement @'
I am applying for ABB's Portfolio / Industry Management and Business Development opportunity because it requires the exact combination I have built throughout my career: mission-critical engineering credibility, customer-facing technical judgement, and the ability to translate complex technical requirements into portfolio strategy, solution mapping and business value. This role is not simply about product explanation; it is about understanding how marine, defense, energy and infrastructure customers make investment decisions when uptime, safety, compliance, lifecycle cost and operational risk are all connected.
'@ @'
I am applying for ABB's Portfolio and Industry Manager role because it requires the exact combination I have built throughout my career: mission-critical engineering credibility, customer-facing technical judgement, and the ability to translate complex technical requirements into offering strategy, portfolio strategy, go-to-market direction and business value. This role is not simply about product explanation; it is about understanding how customers make investment decisions when uptime, safety, compliance, lifecycle cost, time-to-market, product profitability and operational risk are all connected.
'@

Add-Replacement @'
For more than 21 years, I have worked in high-reliability engineering environments covering naval shipbuilding, defence electronics, satellite communication systems and customer technical delivery. My strongest foundation is the 15 years I spent at Daeyang Electric, supporting ROK Navy ship and submarine programs including FFX Batch-II and Jangbogo/KSS-III. In those projects, requirements, interfaces, test evidence, FAT/SAT, sea trials, acceptance records and after-delivery support had to be managed with discipline, because customer trust depended on clear evidence and reliable follow-up.
'@ @'
For more than 21 years, I have worked in high-reliability engineering environments covering naval shipbuilding, defence electronics, satellite communication systems and customer technical delivery. My strongest foundation is the 15 years I spent at Daeyang Electric, supporting ROK Navy ship and submarine programs including FFX Batch-II and Jangbogo/KSS-III. In those projects, requirements, interfaces, test evidence, FAT/SAT, sea trials, acceptance records and after-delivery support had to be managed with discipline. In parallel, I supported Sales HQ whenever new special-vessel opportunities arose by structuring customer needs, estimating development cost, manpower, materials and outsourcing scope, and preparing technical proposal inputs.
'@

Add-Replacement @'
That experience is the reason ABB is highly attractive to me. Marine and infrastructure customers do not purchase automation, control, electrification or lifecycle services in isolation. They are looking for safer operations, higher uptime, stronger system integration, faster issue recovery, lower lifecycle risk, improved energy efficiency and trusted local support. ABB's Marine & Ports, Industrial Automation & Control, and Power & Energy Infrastructure portfolio is positioned to turn these operational problems into measurable business outcomes.
'@ @'
That experience is the reason ABB is highly attractive to me. Industrial, machine automation, marine and infrastructure customers do not purchase automation, control, electrification or lifecycle services in isolation. They are looking for safer operations, higher uptime, stronger system integration, faster issue recovery, lower lifecycle risk, improved energy efficiency, shorter time-to-market and trusted local support. ABB's Machine Automation and Industrial Automation portfolio can turn these operational problems into measurable business outcomes when the offering strategy and segment strategy are grounded in real customer constraints.
'@

Add-Replacement @'
My core strength is the ability to stand between the customer, engineering teams and business stakeholders, then convert technical reality into a clear action plan. In shipbuilding and defense projects, I learned that the real issue is often not a single defect or requirement. It is the lack of a shared view across interfaces, operating conditions, schedule risk, test evidence, acceptance status and downstream lifecycle impact. This is exactly where a strong technical strategy and portfolio manager must create value.
'@ @'
My core strength is the ability to stand between the customer, engineering teams and business stakeholders, then convert technical reality into a clear action plan. In shipbuilding and defense projects, I learned that the real issue is often not a single defect or requirement. It is the lack of a shared view across interfaces, operating conditions, schedule risk, cost assumptions, test evidence, acceptance status and downstream lifecycle impact. This is exactly where a strong portfolio and industry manager must create value by structuring complex problems, eliciting unspoken technical requirements, and translating them into offering priorities, sales enablement and portfolio feedback.
'@

Add-Replacement @'
At Daeyang Electric, I converted Navy and shipyard requirements into system architecture, equipment configuration, ICDs, test plans, test procedures, acceptance criteria and close-out documents. I also supported installation, commissioning, FAT, SAT, sea trial and customer acceptance activities with Navy users, shipyards, inspectors, subcontractors, production, quality and engineering teams. This gives me practical credibility when discussing real shipyard pain points, not only software concepts.
'@ @'
At Daeyang Electric, I converted Navy and shipyard requirements into system architecture, equipment configuration, ICDs, test plans, test procedures, acceptance criteria and close-out documents. I also supported Sales HQ on each new special-vessel opportunity by clarifying customer requirements, defining technical configuration, estimating development cost, manpower, material and outsourcing scope, reviewing quotation assumptions, and preparing proposal materials. In the Indonesia submarine entertainment/broadcasting project, this technical-sales support contributed to an approximately KRW 1.2B business win.
'@

Add-Replacement @'
I can connect that experience to ABB-relevant customer environments in Korea and APAC: shipyards, naval programs, ports, onboard systems, communication/control infrastructure, power and energy interfaces, and high-availability industrial operations. This process view is important because many customer conversations begin with symptoms such as recurring field issues, unclear root cause ownership, delayed acceptance, interface instability, lifecycle service gaps, energy-efficiency pressure or fragmented operational data.
'@ @'
I can connect that experience to ABB-relevant customer environments in Korea and APAC: machine automation customers, shipyards, naval programs, ports, onboard systems, communication/control infrastructure, power and energy interfaces, and high-availability industrial operations. This process view is important because many customer conversations begin with symptoms such as recurring field issues, unclear root-cause ownership, delayed acceptance, interface instability, lifecycle service gaps, cost pressure, product-profitability constraints or fragmented operational data.
'@

Add-Replacement @'
Commercially, this background fits the technical value-selling and portfolio-positioning side of the role. I can support Account Managers, Product/Portfolio teams and Pre-Sales or Service specialists by structuring discovery questions, clarifying customer pain points, preparing value-proposition inputs, shaping proof-of-value storylines and turning technical discussion into CRM-ready next actions. My contribution is strongest where customer needs must be interpreted accurately before a solution can be positioned credibly.
'@ @'
Commercially, this background fits the technical value-selling and portfolio-positioning side of the role. I can support Account Managers, Product/Portfolio teams and Pre-Sales or Service specialists by structuring discovery questions, clarifying customer pain points, preparing value-proposition inputs, shaping proof-of-value storylines, supporting sales tools and turning technical discussion into CRM-ready next actions. My contribution is strongest where customer needs must be interpreted accurately before an offering, product-line portfolio or GTM model can be positioned credibly.
'@

Add-Replacement @'
I have studied marine and industrial portfolio strategy from the viewpoints of product position, competitor strengths, customer pain points and future go-to-market logic, and consolidated that work into my own executive strategy materials and KPI-based value model. That work strengthened my motivation to move from pure engineering delivery into technical strategy and business development, because I can see clear strategic opportunities in marine, ports, industrial automation and power/energy infrastructure customers.
'@ @'
I have studied marine and industrial portfolio strategy from the viewpoints of product position, competitor strengths, customer pain points, market intelligence and future go-to-market logic, and consolidated that work into my own executive strategy materials and KPI-based value model. That work strengthened my motivation to move from pure engineering delivery into portfolio strategy and business development, because I can see clear strategic opportunities in machine automation, industrial automation, marine, ports and power/energy infrastructure customers.
'@

Add-Replacement @'
My practical perspective is that ABB can win when its portfolio is positioned around customer outcomes rather than isolated hardware or software functions. Marine and infrastructure customers care about mission continuity, automation/control reliability, energy efficiency, maintainability, cybersecurity-aware connectivity, lifecycle service response and local execution credibility. My background allows me to connect ABB's portfolio language to real customer conditions: shipboard interfaces, commissioning constraints, acceptance evidence, field issue closure, after-delivery support and executive-level business risk.
'@ @'
My practical perspective is that ABB can win when its offering portfolio is positioned around customer outcomes rather than isolated hardware or software functions. Machine automation, marine and infrastructure customers care about mission continuity, automation/control reliability, time-to-market, energy efficiency, maintainability, cybersecurity-aware connectivity, lifecycle service response and local execution credibility. My background allows me to connect ABB's portfolio language to real customer conditions: control interfaces, commissioning constraints, acceptance evidence, field issue closure, after-delivery support, cost assumptions and executive-level business risk.
'@

Add-Replacement @'
For Korea/APAC customers, I would position ABB as a trusted technology partner that connects reliable automation, control, electrification, digital lifecycle insight and field-service credibility. This means mapping customer requirements to the right portfolio themes, explaining integration and lifecycle value in business language, and feeding customer field insights back into portfolio optimization and GTM priorities.
'@ @'
For Korea/APAC customers, I would position ABB as a trusted technology partner that connects reliable automation, control, electrification, digital lifecycle insight and field-service credibility. This means mapping customer requirements to the right offering themes, explaining integration and lifecycle value in business language, and feeding customer field insights back into portfolio optimization, roadmap planning, lifecycle management and GTM priorities.
'@

Add-Replacement @'
After joining, my first priority would be to align quickly with ABB's assigned portfolio, Korea/APAC account priorities, pipeline discipline, partner and channel structure, regional marketing activities and the collaboration model between Account Managers, Portfolio/Product teams, Service teams, Business Development leadership and technical specialists. I understand that a portfolio or industry manager must create leverage for the sales organization, not work as an isolated technical specialist.
'@ @'
After joining, my first priority would be to align quickly with ABB's assigned Machine Automation portfolio, Korea/Japan account priorities, pipeline discipline, partner and channel structure, regional marketing activities and the collaboration model between Account Managers, Portfolio/Product teams, Service teams, Business Development leadership and technical specialists. I understand that a portfolio or industry manager must create leverage for the sales organization by connecting offering strategy, segment strategy, sales enablement and product-line execution.
'@

Add-Replacement @'
My contribution would begin with customer discovery. I would help identify and organize marine, port, defense, energy and industrial pain points around system integration, automation/control reliability, power-interface resilience, field issue recurrence, commissioning risk, lifecycle service readiness, energy efficiency and operational data ownership. These pain points can then be converted into value hypotheses, discovery questions, benchmark assumptions, benefit estimates, proof-of-value themes and clear follow-up actions.
'@ @'
My contribution would begin with customer discovery and market intelligence. I would help identify and organize machine automation, marine, port, defense, energy and industrial pain points around system integration, automation/control reliability, field issue recurrence, commissioning risk, lifecycle service readiness, energy efficiency, operational data ownership and product profitability. These pain points can then be converted into value hypotheses, discovery questions, benchmark assumptions, benefit estimates, sales tools, proof-of-value themes and clear follow-up actions.
'@

Add-Replacement @'
I would also support executive-level communication by translating technical detail into business language: cost, schedule, risk, uptime, readiness, compliance, productivity, energy efficiency and sustainability. My experience with ICDs, test plans, VCRM-style evidence packages, ECO/BOM changes, WBS schedules, cost estimates, defect logs and acceptance documents gives me the working discipline to make customer discussions specific and credible.
'@ @'
I would also support executive-level communication by translating technical detail into business language: cost, schedule, risk, uptime, readiness, compliance, productivity, time-to-market, product profitability, energy efficiency and sustainability. My experience with ICDs, test plans, VCRM-style evidence packages, ECO/BOM changes, WBS schedules, cost estimates, quotation assumptions, defect logs and acceptance documents gives me the working discipline to make customer discussions specific and credible.
'@

Add-Replacement @'
My goal is to help ABB grow meaningful Marine & Ports, Industrial Automation and Power/Energy Infrastructure opportunities in Korea/APAC by combining mission-critical domain credibility, system-integration understanding, portfolio-positioning discipline and evidence-based customer engagement, while making ABB's value clear to customers who need practical answers to real operational and lifecycle problems.
'@ @'
My goal is to help ABB grow meaningful Machine Automation, Industrial Automation, Marine & Ports and Power/Energy Infrastructure opportunities in Korea/APAC by combining mission-critical domain credibility, system-integration understanding, offering/portfolio-positioning discipline and evidence-based customer engagement, while making ABB's value clear to customers who need practical answers to real operational, profitability and lifecycle problems.
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
try { $word.DisplayAlerts = 0 } catch {}
try { $word.EventsEnabled = $false } catch {}
try { $word.AutomationSecurity = 3 } catch {}
$doc = $null

try {
    $confirmConversions = $false
    $readOnly = $false
    $addToRecentFiles = $false
    $doc = $word.Documents.Open([ref]$OutPath, [ref]$confirmConversions, [ref]$readOnly, [ref]$addToRecentFiles)

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

    Write-Host "Saved DOCX: $OutPath"
    Write-Host "Saved PDF : $PdfPath"
    Write-Host "Pages     : $($doc.ComputeStatistics(2))"
    Write-Host "Words     : $($doc.ComputeStatistics(0))"
    Write-Host "Hits      : $hits"
    Write-Host "Misses    : $($misses.Count)"
    foreach ($m in $misses) {
        Write-Host "MISS: $m"
    }
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
