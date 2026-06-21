$ErrorActionPreference = 'Stop'

$src = 'D:\이력서\ABB - Portfolio and Industry Manager\Joonha_Lee_ABB_Portfolio_Industry_Manager_Marine_Ports_Industrial_Automation_Resume_JD_SalesSupport_Revised.docx'
$outDir = 'D:\이력서\ABB - Portfolio and Industry Manager'
$outDocx = Join-Path $outDir 'Joonha_Lee_ABB_Portfolio_Industry_Manager_Marine_Ports_Industrial_Automation_Resume_JD_SalesSupport_KeywordEnhanced.docx'
$outPdf = [System.IO.Path]::ChangeExtension($outDocx, '.pdf')

if (-not (Test-Path -LiteralPath $src)) {
    throw "Source file not found: $src"
}

Copy-Item -LiteralPath $src -Destination $outDocx -Force

function Normalize-Text([string]$text) {
    return (($text -replace "[`r`n`a\x07]", "") -replace "\s+", " ").Trim()
}

function Replace-ParagraphText($doc, [string]$oldText, [string]$newText) {
    $target = Normalize-Text $oldText
    foreach ($p in $doc.Paragraphs) {
        $current = Normalize-Text $p.Range.Text
        if ($current -eq $target) {
            $r = $p.Range.Duplicate
            if ($r.End -gt $r.Start) {
                $r.End = $r.End - 1
            }
            $r.Text = $newText
            $p.Range.Font.Bold = [int]0
            return $true
        }
    }
    return $false
}

$replacements = @(
    @{
        Old = 'Managed marine system integration, technical documentation, FAT/SAT, sea trial support, customer acceptance, field issue response and coordination with Navy, shipyards, inspectors, subcontractors, production, quality, engineering and sales teams. For every new special-vessel sales opportunity during Daeyang Electric tenure, supported Sales HQ by translating customer requirements into technical configuration, scope, development cost, manpower, material/outsourcing cost, quotation logic and proposal content. Mapped this delivery and sales-support experience to ABB-relevant value themes: offering strategy, time-to-market, sales enablement, product profitability, lifecycle service readiness and trusted customer acceptance.'
        New = 'Managed marine system integration, technical documentation, FAT/SAT, sea trial support, customer acceptance, field issue response and coordination with Navy, shipyards, inspectors, subcontractors, development, quality, production, engineering and sales teams. In each new special-vessel sales opportunity, acted as the middle translator between Sales HQ, customers and execution teams: organizing complex problems, drawing out unspoken technical requirements, converting them into technical configuration, scope, development cost, manpower, material/outsourcing cost, quotation logic and proposal content, and mapping the result to ABB-relevant value themes such as offering strategy, time-to-market, sales enablement, product profitability, lifecycle service readiness and trusted customer acceptance.'
    },
    @{
        Old = 'Analysed customer requirements from the Navy, shipyards, and defence authorities, including requirements that customers could not fully articulate, and translated them into system architecture, network configuration, equipment specifications, ICDs, technical proposals, cost-estimation inputs and execution plans.'
        New = 'Analysed customer requirements from the Navy, shipyards, and defence authorities, including operational needs and technical constraints that customers could not fully articulate. Converted ambiguous statements, site constraints and stakeholder concerns into system architecture, network configuration, equipment specifications, ICDs, technical proposals, cost-estimation inputs and execution plans.'
    },
    @{
        Old = 'Converted customer needs into implementable system configurations, quotation assumptions, milestone-based execution plans and sales-ready value propositions.'
        New = 'Converted customer needs into implementable system configurations, quotation assumptions, milestone-based execution plans and sales-ready value propositions, turning engineering feasibility into business opportunity and giving Sales HQ a practical basis for customer negotiation, pricing logic and proposal differentiation.'
    },
    @{
        Old = 'Identified schedule and cost risks early and aligned sales, purchasing, production, quality, engineering and subcontractor actions to keep projects within quotation, delivery and acceptance targets.'
        New = 'Identified schedule, cost, quality and delivery risks early and aligned development, quality, production, sales, purchasing and subcontractor actions to keep projects within quotation, delivery and acceptance targets. This cross-functional translator role helped each group understand not only its own task, but also the customer impact, cost impact and acceptance risk behind the decision.'
    },
    @{
        Old = 'Led and supported FAT, SAT, commissioning, sea trial, and customer acceptance activities with the Navy, shipbuilders, inspectors, and subcontractors, feeding site learnings back into recurrence-prevention structures and future quotation assumptions.'
        New = 'Led and supported FAT, SAT, commissioning, sea trial, and customer acceptance activities with the Navy, shipbuilders, inspectors, and subcontractors. When field issues occurred, analysed symptoms, reproduction conditions, interface ownership and evidence data, then fed site learnings back into recurrence-prevention structures, design/documentation improvements and future quotation assumptions.'
    },
    @{
        Old = 'Minimized project closing risk by resolving issues through root-cause analysis, corrective action and recurrence-prevention structures rather than responsibility disputes.'
        New = 'Minimized project closing risk by resolving issues through root-cause analysis, corrective action, re-verification criteria and recurrence-prevention structures rather than responsibility disputes, creating a repeatable issue-closure model that protected customer trust and lifecycle business potential.'
    },
    @{
        Old = 'Contributed to an approximately KRW 1.2B Indonesia submarine entertainment/broadcasting system opportunity by proposing a wireless network architecture, supporting cost estimation, persuading customer/shipyard stakeholders, and linking technology to business opportunity.'
        New = 'Contributed to an approximately KRW 1.2B Indonesia submarine entertainment/broadcasting system opportunity by proposing a wireless network architecture, supporting development/manpower/material/outsourcing cost estimation, persuading customer and shipyard stakeholders, and converting a technical solution into a concrete business opportunity.'
    },
    @{
        Old = 'My core strength is the ability to stand between the customer, engineering teams and business stakeholders, then convert technical reality into a clear action plan. In shipbuilding and defense projects, I learned that the real issue is often not a single defect or requirement. It is the lack of a shared view across interfaces, operating conditions, schedule risk, cost assumptions, test evidence, acceptance status and downstream lifecycle impact. This is exactly where a strong portfolio and industry manager must create value by structuring complex problems, eliciting unspoken technical requirements, and translating them into offering priorities, sales enablement and portfolio feedback.'
        New = 'My core strength is the ability to stand between customers, development, quality, production, sales and business stakeholders, then convert technical reality into a clear action plan. In shipbuilding and defense projects, I learned that the real issue is often not a single defect or requirement. It is the lack of a shared view across interfaces, operating conditions, schedule risk, cost assumptions, test evidence, acceptance status and downstream lifecycle impact. This is exactly where a strong portfolio and industry manager must create value: structuring complex problems, drawing out unspoken technical requirements, translating across internal and customer interfaces, analysing failure causes, building recurrence-prevention structures and converting technology into sales enablement, offering priorities and business opportunities.'
    },
    @{
        Old = 'At Daeyang Electric, I converted Navy and shipyard requirements into system architecture, equipment configuration, ICDs, test plans, test procedures, acceptance criteria and close-out documents. I also supported Sales HQ on each new special-vessel opportunity by clarifying customer requirements, defining technical configuration, estimating development cost, manpower, material and outsourcing scope, reviewing quotation assumptions, and preparing proposal materials. In the Indonesia submarine entertainment/broadcasting project, this technical-sales support contributed to an approximately KRW 1.2B business win.'
        New = 'At Daeyang Electric, I converted Navy and shipyard requirements into system architecture, equipment configuration, ICDs, test plans, test procedures, acceptance criteria and close-out documents. I also supported Sales HQ on each new special-vessel opportunity by clarifying explicit requirements, eliciting unspoken technical needs, defining technical configuration, estimating development cost, manpower, material and outsourcing scope, reviewing quotation assumptions, and preparing proposal materials. In practice, I often served as the translator between customer expectations, engineering feasibility, quality evidence, production constraints and sales commitments. In the Indonesia submarine entertainment/broadcasting project, this technical-sales support contributed to an approximately KRW 1.2B business win.'
    },
    @{
        Old = 'Commercially, this background fits the technical value-selling and portfolio-positioning side of the role. I can support Account Managers, Product/Portfolio teams and Pre-Sales or Service specialists by structuring discovery questions, clarifying customer pain points, preparing value-proposition inputs, shaping proof-of-value storylines, supporting sales tools and turning technical discussion into CRM-ready next actions. My contribution is strongest where customer needs must be interpreted accurately before an offering, product-line portfolio or GTM model can be positioned credibly.'
        New = 'Commercially, this background fits the technical value-selling and portfolio-positioning side of the role. I can support Account Managers, Product/Portfolio teams and Pre-Sales or Service specialists by structuring discovery questions, clarifying customer pain points, preparing value-proposition inputs, shaping proof-of-value storylines, supporting sales tools and turning technical discussion into CRM-ready next actions. My contribution is strongest where complex customer problems must be organized, latent needs must be surfaced, field failures must be translated into recurrence-prevention and lifecycle-service logic, and customer needs must be interpreted accurately before an offering, product-line portfolio or GTM model can be positioned credibly.'
    },
    @{
        Old = 'My contribution would begin with customer discovery and market intelligence. I would help identify and organize machine automation, marine, port, defense, energy and industrial pain points around system integration, automation/control reliability, field issue recurrence, commissioning risk, lifecycle service readiness, energy efficiency, operational data ownership and product profitability. These pain points can then be converted into value hypotheses, discovery questions, benchmark assumptions, benefit estimates, sales tools, proof-of-value themes and clear follow-up actions.'
        New = 'My contribution would begin with customer discovery and market intelligence. I would help identify and organize machine automation, marine, port, defense, energy and industrial pain points around system integration, automation/control reliability, field issue recurrence, commissioning risk, lifecycle service readiness, energy efficiency, operational data ownership and product profitability. By asking the questions customers cannot yet formulate, translating between technical and commercial stakeholders, and converting field issues into recurrence-prevention and lifecycle-value logic, these pain points can be converted into value hypotheses, discovery questions, benchmark assumptions, benefit estimates, sales tools, proof-of-value themes and clear follow-up actions.'
    }
)

$word = New-Object -ComObject Word.Application
$word.Visible = $false

$misses = New-Object System.Collections.Generic.List[string]
$changed = 0

try {
    $doc = $word.Documents.Open($outDocx, $false, $false)
    foreach ($item in $replacements) {
        if (Replace-ParagraphText $doc $item.Old $item.New) {
            $changed++
        } else {
            $misses.Add($item.Old.Substring(0, [Math]::Min(120, $item.Old.Length)))
        }
    }

    $doc.Save()
    $doc.ExportAsFixedFormat($outPdf, 17)

    $pages = $doc.ComputeStatistics(2)
    $words = $doc.Words.Count
    $paras = $doc.Paragraphs.Count
    $tables = $doc.Tables.Count
    $doc.Close([ref]$true)
} finally {
    if ($word.Documents.Count -eq 0) {
        $word.Quit()
    }
}

Write-Host "Output DOCX: $outDocx"
Write-Host "Output PDF : $outPdf"
Write-Host "Replacements applied: $changed / $($replacements.Count)"
Write-Host "Pages=$pages Words=$words Paragraphs=$paras Tables=$tables"
if ($misses.Count -gt 0) {
    Write-Host "Misses:"
    foreach ($m in $misses) {
        Write-Host " - $m"
    }
    exit 2
}
