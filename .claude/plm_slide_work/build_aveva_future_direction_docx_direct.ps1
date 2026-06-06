$ErrorActionPreference = 'Stop'

$workDir = 'C:\Users\namma\.claude\plm_slide_work\word_future_direction'
$proposal = 'D:\이력서\AVEVA - Marine Principal Technical Support & Consultant – PLM SME, Busan\Proposal'
$tempDocx = Join-Path $workDir 'AVEVA_Marine_Future_Direction_Meeting_Materials_EN.docx'
$finalDocx = Join-Path $proposal 'AVEVA_Marine_Future_Direction_Development_Planning_Meeting_Materials_EN.docx'
$stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$backupDir = Join-Path $proposal ("_backup_20260606_word_future_direction_docx_" + $stamp)

if (-not (Test-Path -LiteralPath $workDir)) {
    New-Item -ItemType Directory -Path $workDir | Out-Null
}
if (Test-Path -LiteralPath $tempDocx) {
    Remove-Item -LiteralPath $tempDocx -Force
}

function Set-Font {
    param(
        [object]$Selection,
        [string]$Name = 'Aptos',
        [single]$Size = 10.5,
        [int]$Bold = 0,
        [int]$Color = 0
    )
    $Selection.Font.Name = $Name
    $Selection.Font.Size = $Size
    $Selection.Font.Bold = $Bold
    $Selection.Font.Color = $Color
}

function Add-Para {
    param(
        [object]$Selection,
        [string]$Text,
        [single]$Size = 10.5,
        [int]$Bold = 0,
        [int]$SpaceAfter = 6,
        [int]$Color = 0
    )
    Set-Font -Selection $Selection -Size $Size -Bold $Bold -Color $Color
    $Selection.ParagraphFormat.SpaceAfter = $SpaceAfter
    $Selection.TypeText($Text)
    $Selection.TypeParagraph()
}

function Add-Heading1 {
    param([object]$Selection, [string]$Text)
    Add-Para -Selection $Selection -Text $Text -Size ([single]15) -Bold 1 -SpaceAfter 8 -Color 7024640
}

function Add-Heading2 {
    param([object]$Selection, [string]$Text)
    Add-Para -Selection $Selection -Text $Text -Size ([single]12.5) -Bold 1 -SpaceAfter 5 -Color 5067059
}

function Add-Bullet {
    param([object]$Selection, [string]$Text)
    Add-Para -Selection $Selection -Text ("- " + $Text) -Size ([single]10.2) -Bold 0 -SpaceAfter 3
}

function Add-Rule {
    param([object]$Selection)
    Add-Para -Selection $Selection -Text '________________________________________________________________________________' -Size ([single]7) -Bold 0 -SpaceAfter 6 -Color 12632256
}

$word = New-Object -ComObject Word.Application
$doc = $null
try {
    $word.Visible = $false
    $doc = $word.Documents.Add()
    $doc.PageSetup.TopMargin = 54
    $doc.PageSetup.BottomMargin = 54
    $doc.PageSetup.LeftMargin = 54
    $doc.PageSetup.RightMargin = 54
    $doc.PageSetup.Orientation = 0

    $sel = $word.Selection

    Add-Para -Selection $sel -Text 'AVEVA Marine Development Future Direction' -Size ([single]20) -Bold 1 -SpaceAfter 4 -Color 7024640
    Add-Para -Selection $sel -Text 'Final English Meeting Materials for Development and Planning Teams' -Size ([single]13) -Bold 1 -SpaceAfter 10 -Color 5067059
    Add-Para -Selection $sel -Text 'Prepared for AVEVA Marine PLM development, product planning and stakeholder alignment meeting.' -Size ([single]10.5) -Bold 0 -SpaceAfter 3
    Add-Para -Selection $sel -Text 'Date: 6 June 2026' -Size ([single]10.5) -Bold 0 -SpaceAfter 12
    Add-Rule -Selection $sel
    Add-Para -Selection $sel -Text 'Purpose: clearly explain and persuade the audience on the future development direction for AVEVA Marine: an open, governed, evidence-driven control plane that connects shipbuilding design, configuration, certified assurance, approval and operations feedback.' -Size ([single]11) -Bold 1 -SpaceAfter 10

    Add-Heading1 -Selection $sel -Text '1. Executive Summary'
    Add-Para -Selection $sel -Text 'Core message: AVEVA Marine should become the open, governed decision control plane that lets shipyards see, explain, approve and replay every major engineering state transition across design, assurance and operations.' -Size ([single]11) -Bold 1 -SpaceAfter 8
    Add-Para -Selection $sel -Text 'The future development direction should not be framed as another isolated engineering tool, a standalone dashboard or a replacement for specialist solvers. The stronger position is to make AVEVA Marine the trusted layer above specialist tools: the layer that owns configuration, impact analysis, evidence, approval, auditability and lifecycle learning.' -Size ([single]10.5) -Bold 0 -SpaceAfter 6
    Add-Para -Selection $sel -Text 'Decision requested: approve the Phase 2 configuration core plus one bounded Continuous Naval Assurance pilot. Use the 26-week delivery WBS and objective KPI gates as the acceptance model.' -Size ([single]10.8) -Bold 1 -SpaceAfter 10

    Add-Heading1 -Selection $sel -Text '2. Why the Future Direction Must Change'
    Add-Para -Selection $sel -Text 'Marine development is moving beyond single-tool productivity. Shipyards need confidence that every engineering decision can be traced from requirement to design revision, BOM/effectivity, affected scenario, solver evidence, approval and operations feedback. Today, that loop is fragmented across authoring systems, PLM systems, calculation tools, handover repositories and operations systems.' -Size ([single]10.5) -Bold 0 -SpaceAfter 6
    Add-Bullet -Selection $sel -Text 'Current gap: no vendor owns the complete change-to-assurance-to-operations loop.'
    Add-Bullet -Selection $sel -Text 'Consequence: teams rely on manual reconciliation, local knowledge and disconnected approvals.'
    Add-Bullet -Selection $sel -Text 'Future AVEVA response: own the cross-domain decision control plane, not every specialist calculation.'
    Add-Bullet -Selection $sel -Text 'AI boundary: use AI for diagnosis, recommendations and gate holds only; humans and class authorities retain release authority.'

    Add-Heading1 -Selection $sel -Text '3. Target Product Position'
    Add-Para -Selection $sel -Text 'The proposed future product category is an Open AVEVA Marine Decision Control Plane. It does not remove AVEVA E3D, Marine, Hull, Draw, NAPA or other approved tools. Instead, it makes their outputs governable and auditable across the lifecycle.' -Size ([single]10.5) -Bold 0 -SpaceAfter 6
    Add-Heading2 -Selection $sel -Text 'Configuration'
    Add-Bullet -Selection $sel -Text 'Own BOM, effectivity, baseline, revision and affected-object relationships.'
    Add-Bullet -Selection $sel -Text 'Customer value: the shipyard can understand which hull, block, option, date or operational state is affected.'
    Add-Heading2 -Selection $sel -Text 'Impact and Evidence'
    Add-Bullet -Selection $sel -Text 'Link requirement, design, BOM, solver scenario, diagnostics, approval and audit history.'
    Add-Bullet -Selection $sel -Text 'Customer value: class/customer reviewers can replay the evidence behind the decision.'
    Add-Heading2 -Selection $sel -Text 'Approval and Operations Learning'
    Add-Bullet -Selection $sel -Text 'Control evaluate, review, hold, promote and override states with role-gated authority.'
    Add-Bullet -Selection $sel -Text 'Feed AIM, PI, CONNECT and operations calibration back into future engineering decisions.'

    Add-Heading1 -Selection $sel -Text '4. Architecture Direction'
    Add-Para -Selection $sel -Text 'The architecture must be explained carefully because it is a major feasibility point for development teams. The correct boundary is hybrid.' -Size ([single]10.5) -Bold 0 -SpaceAfter 6
    Add-Bullet -Selection $sel -Text 'Windows Authoring Tier: AVEVA E3D, Marine, Hull and Draw remain Windows-centered authoring environments.'
    Add-Bullet -Selection $sel -Text 'Local Agent / Plugin Bridge: authoring tools connect through Local Agent, Plugin or PML.NET bridge patterns.'
    Add-Bullet -Selection $sel -Text 'Linux Control Plane: API, SQL, AI Gate, evidence services, integration gateway and dashboards run on Linux infrastructure.'
    Add-Bullet -Selection $sel -Text 'Cloud/CONNECT Integration: AVEVA CONNECT, AIM, PI and related services are integrated through governed adapters.'
    Add-Bullet -Selection $sel -Text 'Certified Solver Federation: NAPA and other approved solvers remain solver-of-record systems; AVEVA records and governs the evidence.'
    Add-Para -Selection $sel -Text 'Implementation boundary to protect: do not assume E3D/Marine/Hull is hosted directly on Linux. Linux hosts the control plane. Windows authoring tools connect through agents, plugins or callback receivers.' -Size ([single]10.5) -Bold 1 -SpaceAfter 8

    Add-Heading1 -Selection $sel -Text '5. AI and Automation Governance'
    Add-Para -Selection $sel -Text 'AI should be positioned as a controlled assistant, not as a certification authority. This makes the proposal more credible to engineers, planners, class reviewers and security stakeholders.' -Size ([single]10.5) -Bold 0 -SpaceAfter 6
    Add-Bullet -Selection $sel -Text 'Rule-based AI Gate: immediate deterministic validation of metadata, drawing standards, VCRM coverage, missing evidence and approval readiness.'
    Add-Bullet -Selection $sel -Text 'On-prem RAG: semantic search over drawings, specifications, change history, evidence packs and asset context using local data.'
    Add-Bullet -Selection $sel -Text 'AVEVA CONNECT Industrial AI: secure Q&A and decision support over connected engineering and asset data when governance is ready.'
    Add-Bullet -Selection $sel -Text 'Unified Engineering / Design AI: routing suggestions, clash alternatives and simulation acceleration, always followed by engineer and class validation.'
    Add-Para -Selection $sel -Text 'Control principle: AI diagnoses and recommends; rule guardrails hold gates; certified solvers calculate; humans and class authorities release.' -Size ([single]10.5) -Bold 1 -SpaceAfter 8

    Add-Heading1 -Selection $sel -Text '6. Continuous Naval Assurance Pilot'
    Add-Para -Selection $sel -Text 'The pilot should be deliberately narrow. It should prove one closed, auditable loop rather than attempt enterprise-wide PLM replacement.' -Size ([single]10.5) -Bold 0 -SpaceAfter 6
    Add-Bullet -Selection $sel -Text 'E3D/Hull event: a design change or revision event is emitted through a Local Agent or Plugin.'
    Add-Bullet -Selection $sel -Text 'Configuration impact: the control plane resolves BOM, effectivity, baseline and affected scenarios.'
    Add-Bullet -Selection $sel -Text 'Solver evidence: NAPA or another approved solver runs against the defined scenario.'
    Add-Bullet -Selection $sel -Text 'AI/rule gate: rules and AI assistance rank issues, missing evidence and review concerns.'
    Add-Bullet -Selection $sel -Text 'Approval and promote: engineer/class authority approves or holds the state.'
    Add-Bullet -Selection $sel -Text 'Operations feedback: relevant operational context is linked back to future decisions.'

    Add-Heading1 -Selection $sel -Text '7. What Development Teams Must Hear'
    Add-Heading2 -Selection $sel -Text 'Build'
    Add-Bullet -Selection $sel -Text 'OpenAPI contracts for objects, BOM, effectivity, baseline, change impact and assurance evidence.'
    Add-Bullet -Selection $sel -Text 'Local Agent or Plugin callbacks from Windows authoring tools.'
    Add-Bullet -Selection $sel -Text 'AI Gate provider interface: rules first, RAG next, CONNECT later.'
    Add-Heading2 -Selection $sel -Text 'Validate'
    Add-Bullet -Selection $sel -Text 'Seed truth cases for hull, block, compartment, weight, loading case and affected scenarios.'
    Add-Bullet -Selection $sel -Text 'E2E tests for evaluate -> review -> promote and rollback.'
    Add-Bullet -Selection $sel -Text 'KPI gates for contract coverage, accuracy, provenance, response time and security.'
    Add-Heading2 -Selection $sel -Text 'Avoid'
    Add-Bullet -Selection $sel -Text 'No direct Linux hosting assumption for E3D/Marine/Hull.'
    Add-Bullet -Selection $sel -Text 'No big-bang replacement of certified solvers.'
    Add-Bullet -Selection $sel -Text 'No automatic AI release authority.'
    Add-Bullet -Selection $sel -Text 'No unbounded process customization without approval gates.'
    Add-Para -Selection $sel -Text 'Development success statement: a small number of interfaces become reliable enough that planning, engineering and class reviewers can trust the decision state.' -Size ([single]10.5) -Bold 1 -SpaceAfter 8

    Add-Heading1 -Selection $sel -Text '8. What Planning Teams Must Hear'
    Add-Bullet -Selection $sel -Text 'Product position: AVEVA owns the control plane above specialist tools. The customer keeps tool choice, but AVEVA owns lifecycle decision, evidence and approval orchestration.'
    Add-Bullet -Selection $sel -Text 'Customer value: fewer blind handovers, faster change impact review, visible assurance evidence, explainable approvals and replayable history from design to operations.'
    Add-Bullet -Selection $sel -Text 'Roadmap discipline: a 26-week gate model prevents wishful scope. Each capability must pass measurable acceptance before promotion.'
    Add-Bullet -Selection $sel -Text 'Decision request: approve the bounded Phase 2 + Continuous Naval Assurance increment and agree owners, KPI thresholds, assumption boundaries and pilot scenario.'
    Add-Para -Selection $sel -Text 'Planning success statement: the future product direction becomes easy to explain: AVEVA Marine becomes the trusted system for seeing, approving and learning from every major shipbuilding change.' -Size ([single]10.5) -Bold 1 -SpaceAfter 8

    Add-Heading1 -Selection $sel -Text '9. Delivery Roadmap'
    Add-Bullet -Selection $sel -Text '0-2 weeks: readiness, backlog, source data, roles and test environment.'
    Add-Bullet -Selection $sel -Text '2-6 weeks: Phase 1 hardening, OIDC, dashboard, CI, seed truth and current gates.'
    Add-Bullet -Selection $sel -Text '6-12 weeks: configuration core, BOM, effectivity, baseline and APIs.'
    Add-Bullet -Selection $sel -Text '12-16 weeks: change impact, ECR/ECO/MOC, impact graph and promote flow.'
    Add-Bullet -Selection $sel -Text '16-22 weeks: naval assurance, E3D/Hull event, solver adapters and evidence pack.'
    Add-Bullet -Selection $sel -Text '22-26 weeks: UAT, adoption, class/customer review and pilot closeout.'

    Add-Heading1 -Selection $sel -Text '10. KPI Acceptance Model'
    Add-Bullet -Selection $sel -Text 'API contract coverage: 100%. Every committed capability is visible, testable and integrable.'
    Add-Bullet -Selection $sel -Text 'E2E gate pass rate: >=95%. Promote/rollback behavior is reliable under CI.'
    Add-Bullet -Selection $sel -Text 'BOM/effectivity accuracy: >=99%. Configuration decisions are trusted.'
    Add-Bullet -Selection $sel -Text 'Impact graph recall: >=95%. Affected objects and naval scenarios are not missed.'
    Add-Bullet -Selection $sel -Text 'Solver-run provenance: 100%. Evidence is replayable and defensible.'
    Add-Bullet -Selection $sel -Text 'Evidence-pack completeness: 100%. Approval has the required rule trace and affected-object context.'
    Add-Bullet -Selection $sel -Text 'AI diagnostic precision: >=85%. AI explanations are useful without becoming release authority.'
    Add-Bullet -Selection $sel -Text 'Dashboard/approval response: <=5 seconds. Reviewers can use the system in real approval workflows.'

    Add-Heading1 -Selection $sel -Text '11. Risks and Mitigations'
    Add-Bullet -Selection $sel -Text 'Scope creep: pilot one E3D/Hull scenario and federate solvers instead of rewriting them.'
    Add-Bullet -Selection $sel -Text 'Solver or class acceptance: keep certified solver as calculation truth; preserve provenance and human/class release.'
    Add-Bullet -Selection $sel -Text 'Data/model mapping complexity: use seeded truth cases for hull, compartment, weight, loading case and rule links.'
    Add-Bullet -Selection $sel -Text 'AI overreach or security concern: limit AI to recommendation and gate-hold support; enforce OIDC/RBAC, signed evidence and audit trail.'
    Add-Bullet -Selection $sel -Text 'Adoption and ownership: use a cross-functional review board with product planning, naval architecture, solution architecture and development.'

    Add-Heading1 -Selection $sel -Text '12. Meeting Facilitation Guide'
    Add-Bullet -Selection $sel -Text 'Slides 3-6: establish current gaps and public product signals.'
    Add-Bullet -Selection $sel -Text 'Slides 7-15: explain the winning product direction, including AI and lightweight routing.'
    Add-Bullet -Selection $sel -Text 'Slides 16-29: translate strategy into architecture, pilot scope, WBS and objective acceptance criteria.'
    Add-Bullet -Selection $sel -Text 'Slides 30-34: re-score the future state, address ownership/risk/evidence and ask for the decision.'
    Add-Bullet -Selection $sel -Text 'Slides 38-41: use when the room needs a more explicit persuasion structure for development or planning stakeholders.'

    Add-Heading1 -Selection $sel -Text '13. Recommended Closing Argument'
    Add-Para -Selection $sel -Text 'The next differentiator for AVEVA Marine is not another isolated tool feature. It is a governed, auditable, operations-aware control plane that lets customers see, explain, approve and replay every major engineering state transition.' -Size ([single]11) -Bold 1 -SpaceAfter 6
    Add-Para -Selection $sel -Text 'Final ask: align on the product direction, approve the bounded pilot scope, assign owners, validate the first customer/class evidence scenario and use KPI gates to decide promotion.' -Size ([single]10.5) -Bold 1 -SpaceAfter 12
    Add-Rule -Selection $sel
    Add-Para -Selection $sel -Text 'Reference deck: Future_Industrial_PLM_Meeting_Deck_EN_V7.pptx. This Word document is written as standalone English meeting material for development and planning teams.' -Size ([single]9) -Bold 0 -SpaceAfter 3 -Color 8421504

    $doc.SaveAs2($tempDocx, 16)
    $doc.Close([ref]$true)
    $doc = $null
} finally {
    if ($doc -ne $null) { $doc.Close([ref]$false) }
    if ($word.Documents.Count -eq 0) { $word.Quit() }
}

New-Item -ItemType Directory -Path $backupDir | Out-Null
if (Test-Path -LiteralPath $finalDocx) {
    Copy-Item -LiteralPath $finalDocx -Destination (Join-Path $backupDir (Split-Path $finalDocx -Leaf)) -Force
}
Copy-Item -LiteralPath $tempDocx -Destination $finalDocx -Force

Get-Item -LiteralPath $tempDocx, $finalDocx | Select-Object FullName, Length, LastWriteTime
