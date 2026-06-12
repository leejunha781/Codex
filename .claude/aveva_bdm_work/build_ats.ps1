# Build condensed 2-page ATS resume (no tables) for AVEVA BDM Engineering (Marine)
$ErrorActionPreference = 'Stop'
$outDocx = 'D:\이력서\AVEVA - Business Development Manager\Joonha_Lee_AVEVA_BDM_Engineering_Marine_Korea_Japan_Seoul_Resume_ATS_2Page.docx'

$word = New-Object -ComObject Word.Application
$word.Visible = $false
try {
    $doc = $word.Documents.Add()
    $ps = $doc.PageSetup
    $ps.TopMargin = $word.CentimetersToPoints(1.5)
    $ps.BottomMargin = $word.CentimetersToPoints(1.5)
    $ps.LeftMargin = $word.CentimetersToPoints(1.8)
    $ps.RightMargin = $word.CentimetersToPoints(1.8)

    $normal = $doc.Styles.Item([int]-1)
    $normal.Font.Name = 'Arial'
    $normal.Font.Size = [single]10
    $normal.ParagraphFormat.LineSpacingRule = [int]0
    $normal.ParagraphFormat.SpaceAfter = [single]2
    $normal.ParagraphFormat.SpaceBefore = [single]0

    $sel = $word.Selection
    $navy = 128 + 0*256 + 64*65536  # BGR int: dark navy-ish (R0 G? ) -> keep simple charcoal
    $charcoal = 64 + 64*256 + 64*65536

    function TypeLine([string]$t) { $script:sel.TypeText($t); $script:sel.TypeParagraph() }

    function Center([string]$t, [single]$size, [int]$bold, [single]$after) {
        $sel.Style = $doc.Styles.Item([int]-1)
        $sel.ParagraphFormat.Alignment = [int]1
        $sel.ParagraphFormat.SpaceAfter = $after
        $sel.Font.Name = 'Arial'; $sel.Font.Size = $size; $sel.Font.Bold = $bold
        TypeLine $t
        $sel.Font.Bold = [int]0; $sel.Font.Size = [single]10
        $sel.ParagraphFormat.Alignment = [int]0
        $sel.ParagraphFormat.SpaceAfter = [single]2
    }

    function Heading([string]$t) {
        $sel.Style = $doc.Styles.Item([int]-1)
        $sel.ParagraphFormat.Alignment = [int]0
        $sel.ParagraphFormat.SpaceBefore = [single]8
        $sel.ParagraphFormat.SpaceAfter = [single]3
        $sel.Font.Name = 'Arial'; $sel.Font.Size = [single]11; $sel.Font.Bold = [int]1
        $sel.Font.Color = $script:charcoal
        $sel.TypeText($t)
        $b = $sel.Paragraphs.Item(1).Range.ParagraphFormat.Borders.Item([int]-3)
        $b.LineStyle = [int]1; $b.LineWidth = [int]6; $b.Color = [int]$script:charcoal
        $sel.TypeParagraph()
        # reset for following body text
        $p2 = $sel.Paragraphs.Item(1).Range.ParagraphFormat.Borders.Item([int]-3)
        $p2.LineStyle = [int]0
        $sel.Font.Bold = [int]0; $sel.Font.Size = [single]10; $sel.Font.Color = [int]0
        $sel.ParagraphFormat.SpaceBefore = [single]0
        $sel.ParagraphFormat.SpaceAfter = [single]2
    }

    function Body([string]$t) {
        $sel.Style = $doc.Styles.Item([int]-1)
        $sel.Font.Name = 'Arial'; $sel.Font.Size = [single]10; $sel.Font.Bold = [int]0; $sel.Font.Color = [int]0
        TypeLine $t
    }

    function Bullet([string]$t) {
        $sel.Style = $doc.Styles.Item([int]-49)
        $sel.Font.Name = 'Arial'; $sel.Font.Size = [single]10; $sel.Font.Bold = [int]0; $sel.Font.Color = [int]0
        $sel.ParagraphFormat.SpaceAfter = [single]2
        TypeLine $t
        $sel.Style = $doc.Styles.Item([int]-1)
    }

    function Company([string]$boldPart, [string]$rest) {
        $sel.Style = $doc.Styles.Item([int]-1)
        $sel.ParagraphFormat.SpaceBefore = [single]5
        $sel.Font.Name = 'Arial'; $sel.Font.Size = [single]10.5; $sel.Font.Bold = [int]1; $sel.Font.Color = [int]0
        $sel.TypeText($boldPart)
        $sel.Font.Bold = [int]0; $sel.Font.Size = [single]10
        $sel.TypeText($rest)
        $sel.TypeParagraph()
        $sel.ParagraphFormat.SpaceBefore = [single]0
    }

    # ---------- CONTENT ----------
    Center 'JOONHA LEE' ([single]15) 1 ([single]1)
    Center 'Marine Engineering & Industrial Software Business Development Professional' ([single]11) 1 ([single]1)
    Center 'Gunpo-si, Gyeonggi-do, Korea (target: Seoul) | +82-10-2731-4581 | leejunha781@gmail.com' ([single]9.5) 0 ([single]1)
    Center 'Target Role: Business Development Manager - Engineering (Marine), Korea / Japan - AVEVA, Seoul' ([single]9.5) 0 ([single]4)

    Heading 'PROFESSIONAL SUMMARY'
    Body 'Marine and industrial engineering professional with 21+ years across ROK Navy shipbuilding programmes (FFX Batch-II, Jangbogo-III), defence electronics and a global LEO satellite terminal programme. Combines shipyard lifecycle credibility - requirements/VCRM, E3D/Marine design data, MTO/eBOM, baseline/effectivity, ECR/ECO, class evidence and handover - with technical value selling: translating marine pain points such as design rework and system integration into value propositions, benchmarks and qualified opportunities. Ready to support AVEVA regional sales teams, Account Managers and Pre-Sales Consultants in Korea/Japan with domain credibility, competitive PLM insight (Siemens, Dassault, Hexagon, PTC, NAPA) and Salesforce CRM-ready pipeline support.'

    Heading 'CORE COMPETENCIES'
    Bullet 'Marine Business Development | Opportunity Qualification | Lead Generation Support | Pipeline Growth (Salesforce CRM)'
    Bullet 'Marine & Shipbuilding Domain Knowledge | ROK Navy Programmes | Shipyard Lifecycle (VCRM, eBOM, ECR/ECO, Baseline/Effectivity)'
    Bullet 'Technical Value Proposition | Benchmarks & Benefit Estimation | Customer Pain-Point Discovery | Executive Communication'
    Bullet 'PLM / Digital Thread / Digital Twin / Smart Shipyard | AVEVA E3D/Marine, Unified Engineering, AIM, PI/CONNECT'
    Bullet 'Competitive Positioning (Siemens, Dassault, Hexagon, PTC, NAPA) | Pre-Sales & Account Team Collaboration'
    Bullet 'FAT / SAT / Commissioning / Sea Trial | Industrial Communication (HART, Modbus, Fieldbus, Industrial Ethernet)'

    Heading 'PROFESSIONAL EXPERIENCE'
    Company 'Intellian Technologies - General Manager, SIT' '  |  Seoul area, Korea  |  Mar 2023 - Jul 2025'
    Bullet 'Led system integration testing, customer technical communication and release-readiness governance for the Eutelsat OneWeb LEO satellite terminal programme, managing 400+ verification items across RF/antenna, control electronics, network and environmental domains.'
    Bullet 'Converted customer requirements, defect history, ECO/BOM changes and verification evidence into traceable release decisions - requirements-to-evidence discipline directly applicable to PLM-based lifecycle value selling.'
    Bullet 'Ran recurring technical reviews with a global customer, turning disputed issues into evidence-based corrective actions and aligned next actions while protecting customer trust within delivery limits.'
    Bullet 'Provided global field and exhibition support (Marlink/STW Germany; Australia), building English customer-facing communication with international partners.'

    Company 'Genohco - Principal Research Engineer / PL-PM' '  |  Korea  |  Aug 2022 - Mar 2023'
    Bullet 'Managed defence/aerospace projects end-to-end from RFP and tender analysis through proposal, development, verification, customer inspection and delivery, including tender conditions and contractual delivery terms.'
    Bullet 'Prepared technical proposals, execution plans, cost estimates and presentation materials, improving bid competitiveness by translating customer requirements into commercially aware solutions.'
    Bullet 'Controlled WBS-based schedule, cost and supplier coordination; maintained traceability from requirements through test evidence to customer acceptance.'

    Company 'Daeyang Electric - Deputy General Manager, Naval System Engineer & PM/PL' '  |  Busan, Korea  |  May 2007 - Apr 2022'
    Bullet 'Delivered ROK Navy ship and submarine communication, broadcasting, internal network and C4I-linked systems over 15 years: requirements analysis, design, ICDs, installation, FAT/SAT, sea trials, customer acceptance and after-delivery support with Navy, shipyards, inspectors and subcontractors.'
    Bullet 'Converted Navy and shipyard requirements and pain points into system architecture, technical proposals and milestone-based execution plans grounded in real shipboard and shipyard constraints.'
    Bullet 'Built first-hand understanding of the standard large-shipyard lifecycle (e.g., HD Hyundai-class yards): requirements/VCRM, basic/detail design, E3D/Marine/Draw, MTO/eBOM, baseline/effectivity, ECR/ECO, production release, class approval evidence and handover.'
    Bullet 'Standardised lifecycle documentation (ICDs, test plans/procedures/reports, acceptance evidence) across subsystem teams; supported overseas submarine field assistance in Indonesia.'

    Company 'Early Career - Hardware / RF Engineer' '  |  Korea  |  Jan 2004 - May 2007'
    Bullet 'Excure Hitron (CCTV/IR camera control boards, certification, production stabilisation), Essetel (RF/mobile hardware verification) and NesLab (embedded/PCB development): built the measurement-based troubleshooting and engineering-to-production discipline later applied to naval and satellite programmes.'

    Heading 'AVEVA / PLM / SMART SHIPYARD PREPARATION'
    Bullet 'Developed working knowledge of AVEVA Marine PLM, Engineering Data, Digital Thread and Smart Shipyard; built an executive competitive-strategy deck positioning AVEVA against Siemens, Dassault, Hexagon, PTC and NAPA, including a KPI-gated delivery and acceptance model (configuration accuracy, change-impact recall, evidence completeness, approval response time).'
    Bullet 'Positioning thesis: AVEVA as the open engineering-to-operations decision layer - configuration, change impact, approval evidence, production readiness and operations feedback above specialist tools - strengthened by AVEVA AI direction (CONNECT Industrial AI Assistant, Generative / Predictive Design AI).'

    Heading 'EDUCATION'
    Bullet 'M.S., Electronic & Electrical Engineering, Pusan National University, Korea (2013 - 2018). Thesis: industrial Ethernet-based three-phase motor control for actuator applications (ARM Cortex-M3).'
    Bullet 'B.S., Control & Instrumentation Engineering, University of Ulsan, Korea (1992 - 2002).'

    Heading 'LANGUAGES & CERTIFICATIONS'
    Bullet 'Korean (native); English (professional working proficiency - global customer programmes; language study in Perth, Australia).'
    Bullet 'Certified NMEA 2000 Networking; Defence PMBOK training; Machine Learning with Python; Python for Data Science, AI & Development.'

    # normalize final paragraph style (avoid trailing styled bullet)
    $last = $doc.Paragraphs.Item($doc.Paragraphs.Count)
    $last.Style = $doc.Styles.Item([int]-1)
    if ((($last.Range.Text) -replace "[\r\n]", '') -eq '') { $last.Range.Delete() | Out-Null }

    $pages = $doc.ComputeStatistics(2)
    Write-Output "Pages: $pages | Words: $($doc.ComputeStatistics(0))"

    $doc.SaveAs([ref]$outDocx)
    $pdf = [System.IO.Path]::ChangeExtension($outDocx, 'pdf')
    $doc.ExportAsFixedFormat($pdf, 17)
    Write-Output "Saved: $outDocx"
    Write-Output "PDF:   $pdf"
    $doc.Close([ref]$true)
} finally {
    if ($word.Documents.Count -eq 0) { $word.Quit() }
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($word)
}
Write-Output 'ATS BUILD DONE'
