# AVEVA BDM resume revision pass — JD alignment + codex compliance + deck reflection
$ErrorActionPreference = 'Stop'

function Normalize([string]$t) {
    $t = $t -replace "[\r\n\a\x0b\x07]", ''
    $t = $t -replace [string][char]0x2018, "'" -replace [string][char]0x2019, "'"
    $t = $t -replace [string][char]0x201C, '"' -replace [string][char]0x201D, '"'
    ($t -replace "\s+", " ").Trim()
}

function Set-ParaText($p, [string]$new) {
    $r = $p.Range.Duplicate
    $t = $r.Text
    $cr = [string][char]13; $bell = [string][char]7
    if ($t.EndsWith($cr + $bell)) { $r.End = $r.End - 2 }
    elseif ($t.EndsWith($cr) -or $t.EndsWith($bell)) { $r.End = $r.End - 1 }
    $r.Text = $new
}

# --- replacement map (normalized find => new full text) ---
$map = [ordered]@{
'Intellian Technologies | Senior Manager, SIT' = 'Intellian Technologies | General Manager, SIT'
'XecureHiteRon' = 'Excure Hitron'
'Exicure Hitron | Assistant Manager, Electronic Technology Team' = 'Excure Hitron | Assistant Manager, Electronic Technology Team'
'Nesslab | Assistant Manager, Development Team' = 'NesLab | Assistant Manager, Development Team'
'Daeyang Electric | Naval System Engineer' = 'Daeyang Electric | Deputy General Manager, Naval System Engineer & PM/PL'
'Leejunha781@gmail.com' = 'leejunha781@gmail.com'
'IELTS 5.0; English language study in Perth, Australia;' = 'Professional working English; English language study in Perth, Australia;'
'Shipbuilding and naval-domain credibility for discovering customer pain points, qualifying Marine Engineering opportunities and supporting regional account teams.' = 'Shipbuilding and naval-domain credibility for discovering customer pain points, qualifying Marine Engineering opportunities, supporting lead generation and keeping pipeline actions Salesforce CRM-ready for regional account teams.'
'Practical understanding of ROKN programmes and the HD Hyundai-standard lifecycle: requirements/VCRM, basic/detail design, E3D/Marine/Draw, MTO/eBOM, baseline/effectivity, ECR/ECO and handover evidence.' = 'Practical understanding of ROKN programmes and the standard large-shipyard lifecycle: requirements/VCRM, basic/detail design, E3D/Marine/Draw, MTO/eBOM, baseline/effectivity, ECR/ECO, class/regulatory evidence and production handover.'
'Built 15 years of naval/shipbuilding delivery credibility across requirements, design, FAT/SAT, sea trial and handover; mapped shipyard pain points to HD Hyundai-standard lifecycle controls.' = 'Built 15 years of naval/shipbuilding delivery credibility across requirements, design, FAT/SAT, sea trial and handover; mapped shipyard pain points to standard large-shipyard lifecycle controls.'
'Interpreted customer requirements accurately during the bidding phase and controlled project execution based on schedule, deliverables, and acceptance criteria during the delivery phase.' = 'Interpreted customer requirements accurately during the bidding phase and controlled project execution based on schedule, deliverables, and acceptance criteria during the delivery phase, including tender conditions and contractual delivery terms.'
'AVEVA Marine PLM / Engineering Data / Digital Thread preparation.' = 'AVEVA Marine PLM / Engineering Data / Digital Thread working knowledge (self-built competitive strategy deck and KPI model).'
'AVEVA E3D/Marine concepts, Unified Engineering, AIM, PI/CONNECT positioning, PLM, Digital Thread, Digital Twin, Smart Shipyard, ERP/MES/PLM boundaries, EPLAN concept, ICD, VCRM, ECO/BOM, WBS, cost estimate, defect log, evidence package, Wireshark, iperf, Linux, VMware and RF measurement equipment.' = 'AVEVA E3D/Marine concepts, Unified Engineering, AIM, PI/CONNECT positioning, PLM, Digital Thread, Digital Twin, Smart Shipyard, ERP/MES/PLM boundaries, EPLAN concept, HART/Modbus/Fieldbus/Industrial Ethernet, Salesforce CRM awareness, ICD, VCRM, ECO/BOM, WBS, cost estimate, defect log, evidence package, Wireshark, iperf, Linux, VMware and RF measurement equipment.'
"When I prepared for the AVEVA Marine PLM consultant role, I did not treat it as short-term interview preparation. I studied AVEVA Marine / PLM from the viewpoints of product position, competitor strengths, customer pain points and future strategy, and developed my own view through the V18 meeting deck. That preparation strengthened my motivation to work for AVEVA because I could see a clear strategic opportunity in the marine industry." = "I have studied AVEVA Marine / PLM in depth - from the viewpoints of product position, competitor strengths, customer pain points and future strategy - and consolidated that work into my own executive strategy deck, including a KPI-gated delivery and acceptance model covering configuration accuracy, change-impact recall, evidence completeness and approval response time. That work strengthened my motivation to join AVEVA because I could see a clear strategic opportunity in the marine industry, and it gives me concrete material for benchmark and benefit-estimation discussions with customers."
"My practical comparison is as follows. Siemens Teamcenter is strong in configuration, BOM and change governance. Dassault 3DEXPERIENCE is strong in collaboration and virtual-twin positioning. Hexagon has strength in asset lifecycle and operations context. PTC is strong in requirements, ALM and traceability patterns. NAPA is trusted in naval architecture and stability calculation. AVEVA's opportunity is different: it can connect engineering truth, lifecycle context and operational feedback through E3D/Marine, Unified Engineering, AIM, PI/CONNECT and open integration." = "My practical comparison is as follows. Siemens Teamcenter is strong in configuration, BOM and change governance. Dassault 3DEXPERIENCE is strong in collaboration and virtual-twin positioning. Hexagon has strength in asset lifecycle and operations context. PTC is strong in requirements, ALM and traceability patterns. NAPA is trusted in naval architecture and stability calculation. AVEVA's opportunity is different: it can connect engineering truth, lifecycle context and operational feedback through E3D/Marine, Unified Engineering, AIM, PI/CONNECT and open integration, strengthened by AVEVA's public AI direction such as the CONNECT Industrial AI Assistant and Generative / Predictive Design AI. The white space is clear: no current vendor owns the complete loop from design change to naval calculation, evidence, approval and operations feedback, and AVEVA is the best-positioned company to own it."
"I can connect that experience to the HD Hyundai-standard shipbuilding lifecycle: requirements/VCRM, basic and detail design, E3D/Marine/Draw, MTO/eBOM, baseline and effectivity control, ECR/ECO, production release and handover evidence. This process view is important for AVEVA because many customer conversations begin with symptoms such as document inconsistency, delayed change review, unclear ownership of engineering data, weak handover evidence or fragmented digital continuity." = "I can connect that experience to the standard large-shipyard lifecycle in Korea and Japan (for example HD Hyundai-class yards): requirements/VCRM, basic and detail design, E3D/Marine/Draw, MTO/eBOM, baseline and effectivity control, ECR/ECO, production release, classification approval evidence and handover. This process view is important for AVEVA because many customer conversations begin with symptoms such as document inconsistency, delayed change review, unclear ownership of engineering data, weak handover evidence or fragmented digital continuity."
}
$msFind = [char]0x2713 + ' M.S. in Electronic & Electrical Engineering, Pusan National University, Korea (Sep 2013 - Aug 2018)'
$msNew  = [char]0x2713 + ' M.S. in Electronic & Electrical Engineering, Pusan National University, Korea (Sep 2013 - Aug 2018) - thesis research on industrial Ethernet-based three-phase motor control for actuator applications (ARM Cortex-M3)'
$map[$msFind] = $msNew

$headline = 'Marine Engineering & Industrial Software Business Development Professional'
$summaryHead = 'Professional Summary'
$summaryBody = 'Marine and industrial engineering professional with 21+ years across ROK Navy shipbuilding programmes (FFX Batch-II, Jangbogo-III), defence electronics and a global LEO satellite terminal programme. Combines shipyard lifecycle credibility - requirements/VCRM, E3D/Marine design data, MTO/eBOM, baseline/effectivity, ECR/ECO, class evidence and handover - with technical value selling: translating marine pain points such as design rework and system integration into value propositions, benchmarks and qualified opportunities. Ready to support AVEVA regional sales teams, Account Managers and Pre-Sales Consultants in Korea/Japan with domain credibility, competitive PLM insight (Siemens, Dassault, Hexagon, PTC, NAPA) and Salesforce CRM-ready pipeline support.'

$files = @(
    'D:\이력서\AVEVA - Business Development Manager\Joonha_Lee_AVEVA_BDM_Engineering_Marine_Korea_Japan_Seoul_Resume_Formal_Tone.docx',
    'D:\이력서\AVEVA - Business Development Manager\Joonha_Lee_AVEVA_BDM_Engineering_Marine_Korea_Japan_Seoul_Resume.docx'
)

$word = New-Object -ComObject Word.Application
$word.Visible = $false
try {
    foreach ($f in $files) {
        Write-Output "=== EDITING: $f"
        $doc = $word.Documents.Open($f, $false, $false)
        try {
            $hits = @{}
            # Pass 1: replacements + fix stray backslash bullet
            foreach ($p in $doc.Paragraphs) {
                $n = Normalize $p.Range.Text
                if ($n -eq '') { continue }
                if ($map.Contains($n)) {
                    Set-ParaText $p $map[$n]
                    $hits[$n] = $true
                } elseif ($n -like 'Built an early foundation in lifecycle thinking by connecting development issues with test, certification and production readiness.*\') {
                    Set-ParaText $p 'Built an early foundation in lifecycle thinking by connecting development issues with test, certification and production readiness.'
                    $hits['stray-backslash'] = $true
                }
            }
            # Pass 2: find anchors fresh
            $biPar = $null; $ccPar = $null; $bodyPar = $null
            foreach ($p in $doc.Paragraphs) {
                $n = Normalize $p.Range.Text
                if ($n -eq 'Basic Information' -and -not $biPar) { $biPar = $p }
                elseif ($n -eq 'Core Competencies' -and -not $ccPar) { $ccPar = $p }
                elseif ($n.StartsWith('I am applying for the Business Development Manager') -and -not $bodyPar) { $bodyPar = $p }
                if ($biPar -and $ccPar -and $bodyPar) { break }
            }
            if (-not ($biPar -and $ccPar -and $bodyPar)) { throw "Anchor not found: BI=$($null -ne $biPar) CC=$($null -ne $ccPar) BODY=$($null -ne $bodyPar)" }
            $srcFont = $bodyPar.Range.Characters.First.Font

            # Insert Professional Summary BEFORE Core Competencies (later in doc first)
            $alreadySummary = $false
            foreach ($p in $doc.Paragraphs) { if ((Normalize $p.Range.Text) -eq $summaryHead) { $alreadySummary = $true; break } }
            if (-not $alreadySummary) {
                $r2 = $ccPar.Range.Duplicate; $r2.Collapse(1)
                $r2.InsertBefore($summaryHead + [char]13 + $summaryBody + [char]13 + [char]13)
                $pars2 = $r2.Paragraphs
                for ($i = 2; $i -le $pars2.Count; $i++) {
                    $pr = $pars2.Item($i).Range
                    $pr.Font.Name = $srcFont.Name
                    $pr.Font.Size = $srcFont.Size
                    $pr.Font.Bold = [int]0
                    try { $pr.Font.Color = $srcFont.Color } catch {}
                    $pr.ParagraphFormat.Alignment = $bodyPar.Format.Alignment
                }
                Write-Output "  + Professional Summary inserted"
            }
            # Insert headline BEFORE Basic Information
            $alreadyHeadline = $false
            foreach ($p in $doc.Paragraphs) { if ((Normalize $p.Range.Text) -eq $headline) { $alreadyHeadline = $true; break } }
            if (-not $alreadyHeadline) {
                $r1 = $biPar.Range.Duplicate; $r1.Collapse(1)
                $r1.InsertBefore($headline + [char]13)
                $r1.Font.Bold = [int]1
                $r1.ParagraphFormat.Alignment = [int]1
                Write-Output "  + Headline inserted"
            }
            Write-Output ("  Replacements applied: {0}/{1}" -f $hits.Count, ($map.Count + 1))
            foreach ($k in $map.Keys) { if (-not $hits.Contains($k)) { Write-Output ("  MISSED: " + $k.Substring(0, [Math]::Min(80, $k.Length))) } }
            if (-not $hits.Contains('stray-backslash')) { Write-Output "  MISSED: stray-backslash bullet" }

            $doc.Save()
            $pdf = [System.IO.Path]::ChangeExtension($f, 'pdf')
            $doc.ExportAsFixedFormat($pdf, 17)
            Write-Output "  Saved + PDF exported: $pdf"
        } finally {
            $doc.Close([ref]$false)
        }
    }
} finally {
    if ($word.Documents.Count -eq 0) { $word.Quit() }
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($word)
}
Write-Output 'EDIT PASS DONE'
