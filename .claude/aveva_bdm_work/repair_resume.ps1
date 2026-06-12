# Repair stray trailing characters left by the End-offset replacement quirk
$ErrorActionPreference = 'Stop'

function VisText([string]$t) {
    ($t -replace "[\r\n\a\x0b\x07]", '').TrimEnd()
}

# exact expected final texts for every replaced/fixed paragraph
$expected = @(
'Intellian Technologies | General Manager, SIT',
'Excure Hitron',
'Excure Hitron | Assistant Manager, Electronic Technology Team',
'NesLab | Assistant Manager, Development Team',
'Daeyang Electric | Deputy General Manager, Naval System Engineer & PM/PL',
'leejunha781@gmail.com',
'Professional working English; English language study in Perth, Australia;',
'Shipbuilding and naval-domain credibility for discovering customer pain points, qualifying Marine Engineering opportunities, supporting lead generation and keeping pipeline actions Salesforce CRM-ready for regional account teams.',
'Practical understanding of ROKN programmes and the standard large-shipyard lifecycle: requirements/VCRM, basic/detail design, E3D/Marine/Draw, MTO/eBOM, baseline/effectivity, ECR/ECO, class/regulatory evidence and production handover.',
'Built 15 years of naval/shipbuilding delivery credibility across requirements, design, FAT/SAT, sea trial and handover; mapped shipyard pain points to standard large-shipyard lifecycle controls.',
'Interpreted customer requirements accurately during the bidding phase and controlled project execution based on schedule, deliverables, and acceptance criteria during the delivery phase, including tender conditions and contractual delivery terms.',
'AVEVA Marine PLM / Engineering Data / Digital Thread working knowledge (self-built competitive strategy deck and KPI model).',
'AVEVA E3D/Marine concepts, Unified Engineering, AIM, PI/CONNECT positioning, PLM, Digital Thread, Digital Twin, Smart Shipyard, ERP/MES/PLM boundaries, EPLAN concept, HART/Modbus/Fieldbus/Industrial Ethernet, Salesforce CRM awareness, ICD, VCRM, ECO/BOM, WBS, cost estimate, defect log, evidence package, Wireshark, iperf, Linux, VMware and RF measurement equipment.',
'Built an early foundation in lifecycle thinking by connecting development issues with test, certification and production readiness.',
([char]0x2713 + ' M.S. in Electronic & Electrical Engineering, Pusan National University, Korea (Sep 2013 - Aug 2018) - thesis research on industrial Ethernet-based three-phase motor control for actuator applications (ARM Cortex-M3)'),
"I have studied AVEVA Marine / PLM in depth - from the viewpoints of product position, competitor strengths, customer pain points and future strategy - and consolidated that work into my own executive strategy deck, including a KPI-gated delivery and acceptance model covering configuration accuracy, change-impact recall, evidence completeness and approval response time. That work strengthened my motivation to join AVEVA because I could see a clear strategic opportunity in the marine industry, and it gives me concrete material for benchmark and benefit-estimation discussions with customers.",
"My practical comparison is as follows. Siemens Teamcenter is strong in configuration, BOM and change governance. Dassault 3DEXPERIENCE is strong in collaboration and virtual-twin positioning. Hexagon has strength in asset lifecycle and operations context. PTC is strong in requirements, ALM and traceability patterns. NAPA is trusted in naval architecture and stability calculation. AVEVA's opportunity is different: it can connect engineering truth, lifecycle context and operational feedback through E3D/Marine, Unified Engineering, AIM, PI/CONNECT and open integration, strengthened by AVEVA's public AI direction such as the CONNECT Industrial AI Assistant and Generative / Predictive Design AI. The white space is clear: no current vendor owns the complete loop from design change to naval calculation, evidence, approval and operations feedback, and AVEVA is the best-positioned company to own it.",
"I can connect that experience to the standard large-shipyard lifecycle in Korea and Japan (for example HD Hyundai-class yards): requirements/VCRM, basic and detail design, E3D/Marine/Draw, MTO/eBOM, baseline and effectivity control, ECR/ECO, production release, classification approval evidence and handover. This process view is important for AVEVA because many customer conversations begin with symptoms such as document inconsistency, delayed change review, unclear ownership of engineering data, weak handover evidence or fragmented digital continuity."
)

function NormApos([string]$t) {
    $t -replace [string][char]0x2018, "'" -replace [string][char]0x2019, "'"
}

$files = @(
    'D:\이력서\AVEVA - Business Development Manager\Joonha_Lee_AVEVA_BDM_Engineering_Marine_Korea_Japan_Seoul_Resume_Formal_Tone.docx',
    'D:\이력서\AVEVA - Business Development Manager\Joonha_Lee_AVEVA_BDM_Engineering_Marine_Korea_Japan_Seoul_Resume.docx'
)

$word = New-Object -ComObject Word.Application
$word.Visible = $false
try {
    foreach ($f in $files) {
        Write-Output "=== REPAIRING: $f"
        $doc = $word.Documents.Open($f, $false, $false)
        try {
            $repaired = 0; $clean = 0
            foreach ($p in $doc.Paragraphs) {
                $vis = NormApos (VisText $p.Range.Text)
                if ($vis -eq '') { continue }
                foreach ($exp in $expected) {
                    $e = NormApos $exp
                    if ($vis -eq $e) { $clean++; break }
                    if ($vis.StartsWith($e) -and ($vis.Length - $e.Length) -le 4) {
                        # delete stray trailing visible chars one at a time (position-safe)
                        $guard = 0
                        while ($guard -lt 6) {
                            $cur = NormApos (VisText $p.Range.Text)
                            if ($cur -eq $e) { break }
                            $chars = $p.Range.Characters
                            $i = $chars.Count
                            while ($i -ge 1) {
                                $ct = $chars.Item($i).Text
                                $code = [int][char]$ct[0]
                                if ($code -ne 13 -and $code -ne 10 -and $code -ne 7 -and $code -ne 11) { break }
                                $i--
                            }
                            if ($i -lt 1) { break }
                            $chars.Item($i).Delete() | Out-Null
                            $guard++
                        }
                        $final = NormApos (VisText $p.Range.Text)
                        if ($final -eq $e) { $repaired++ } else { Write-Output ("  STILL BAD: " + $final.Substring([Math]::Max(0,$final.Length-60))) }
                        break
                    }
                }
            }
            Write-Output "  repaired: $repaired, already clean: $clean (target total: $($expected.Count))"
            # strict verification
            $found = @{}
            foreach ($p in $doc.Paragraphs) {
                $vis = NormApos (VisText $p.Range.Text)
                foreach ($exp in $expected) {
                    if ($vis -eq (NormApos $exp)) { $found[$exp] = $true }
                }
            }
            $missing = @($expected | Where-Object { -not $found.Contains($_) })
            if ($missing.Count -gt 0) {
                foreach ($m in $missing) { Write-Output ("  MISSING EXACT: " + $m.Substring(0, [Math]::Min(70, $m.Length))) }
            } else { Write-Output "  ALL $($expected.Count) paragraphs verified exact." }
            $doc.Save()
            $pdf = [System.IO.Path]::ChangeExtension($f, 'pdf')
            $doc.ExportAsFixedFormat($pdf, 17)
            Write-Output "  Saved + PDF re-exported."
        } finally {
            $doc.Close([ref]$false)
        }
    }
} finally {
    if ($word.Documents.Count -eq 0) { $word.Quit() }
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($word)
}
Write-Output 'REPAIR DONE'
