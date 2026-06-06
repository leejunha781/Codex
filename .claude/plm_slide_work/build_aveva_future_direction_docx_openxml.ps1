$ErrorActionPreference = 'Stop'

$workDir = 'C:\Users\namma\.claude\plm_slide_work\word_future_direction'
$buildDir = Join-Path $workDir 'docx_build'
$tempDocx = Join-Path $workDir 'AVEVA_Marine_Future_Direction_Meeting_Materials_EN.docx'
$proposal = 'D:\이력서\AVEVA - Marine Principal Technical Support & Consultant – PLM SME, Busan\Proposal'
$finalDocx = Join-Path $proposal 'AVEVA_Marine_Future_Direction_Development_Planning_Meeting_Materials_EN.docx'
$finalTxt = Join-Path $proposal 'AVEVA_Marine_Future_Direction_Development_Planning_Meeting_Materials_EN.txt'
$stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$backupDir = Join-Path $proposal ("_backup_20260606_word_future_direction_openxml_" + $stamp)

if (Test-Path -LiteralPath $buildDir) { Remove-Item -LiteralPath $buildDir -Recurse -Force }
if (Test-Path -LiteralPath $tempDocx) { Remove-Item -LiteralPath $tempDocx -Force }
New-Item -ItemType Directory -Path $buildDir | Out-Null
New-Item -ItemType Directory -Path (Join-Path $buildDir '_rels') | Out-Null
New-Item -ItemType Directory -Path (Join-Path $buildDir 'docProps') | Out-Null
New-Item -ItemType Directory -Path (Join-Path $buildDir 'word') | Out-Null
New-Item -ItemType Directory -Path (Join-Path $buildDir 'word\_rels') | Out-Null

function XmlEscape {
    param([string]$Text)
    if ($null -eq $Text) { return '' }
    return [System.Security.SecurityElement]::Escape($Text)
}

function New-Paragraph {
    param(
        [string]$Text,
        [int]$Size = 21,
        [bool]$Bold = $false,
        [string]$Color = '162033',
        [int]$After = 120,
        [int]$Before = 0
    )
    $boldXml = if ($Bold) { '<w:b/>' } else { '' }
    $escaped = XmlEscape $Text
    return "<w:p><w:pPr><w:spacing w:before=`"$Before`" w:after=`"$After`"/></w:pPr><w:r><w:rPr>$boldXml<w:color w:val=`"$Color`"/><w:sz w:val=`"$Size`"/><w:szCs w:val=`"$Size`"/></w:rPr><w:t xml:space=`"preserve`">$escaped</w:t></w:r></w:p>"
}

function New-Heading1 {
    param([string]$Text)
    return New-Paragraph -Text $Text -Size 30 -Bold $true -Color '0B2340' -After 160 -Before 200
}

function New-Heading2 {
    param([string]$Text)
    return New-Paragraph -Text $Text -Size 25 -Bold $true -Color '0D5B72' -After 100 -Before 120
}

function New-Bullet {
    param([string]$Text)
    return New-Paragraph -Text ("- " + $Text) -Size 21 -Bold $false -Color '162033' -After 60
}

function New-Callout {
    param([string]$Text)
    return "<w:p><w:pPr><w:pBdr><w:left w:val=`"single`" w:sz=`"18`" w:space=`"6`" w:color=`"1BCAD6`"/></w:pBdr><w:spacing w:before=`"80`" w:after=`"120`"/></w:pPr><w:r><w:rPr><w:b/><w:color w:val=`"0B2340`"/><w:sz w:val=`"23`"/><w:szCs w:val=`"23`"/></w:rPr><w:t xml:space=`"preserve`">$(XmlEscape $Text)</w:t></w:r></w:p>"
}

$parts = New-Object System.Collections.Generic.List[string]
$plain = New-Object System.Collections.Generic.List[string]

function AddDoc {
    param([string]$Xml, [string]$PlainText = '')
    $parts.Add($Xml)
    if ($PlainText.Length -gt 0) { $plain.Add($PlainText) }
}

AddDoc (New-Paragraph 'AVEVA Marine Development Future Direction' 42 $true '0B2340' 80) 'AVEVA Marine Development Future Direction'
AddDoc (New-Paragraph 'Final English Meeting Materials for Development and Planning Teams' 28 $true '0D5B72' 160) 'Final English Meeting Materials for Development and Planning Teams'
AddDoc (New-Paragraph 'Prepared for AVEVA Marine PLM development, product planning and stakeholder alignment meeting.' 22 $false '162033' 60) 'Prepared for AVEVA Marine PLM development, product planning and stakeholder alignment meeting.'
AddDoc (New-Paragraph 'Date: 6 June 2026' 22 $false '162033' 180) 'Date: 6 June 2026'
AddDoc (New-Callout 'Purpose: clearly explain and persuade the audience on the future development direction for AVEVA Marine: an open, governed, evidence-driven control plane that connects shipbuilding design, configuration, certified assurance, approval and operations feedback.') 'Purpose: clearly explain and persuade the audience on the future development direction for AVEVA Marine.'

AddDoc (New-Heading1 '1. Executive Summary') '1. Executive Summary'
AddDoc (New-Callout 'Core message: AVEVA Marine should become the open, governed decision control plane that lets shipyards see, explain, approve and replay every major engineering state transition across design, assurance and operations.') 'Core message: AVEVA Marine should become the open, governed decision control plane.'
AddDoc (New-Paragraph 'The future development direction should not be framed as another isolated engineering tool, a standalone dashboard or a replacement for specialist solvers. The stronger position is to make AVEVA Marine the trusted layer above specialist tools: the layer that owns configuration, impact analysis, evidence, approval, auditability and lifecycle learning.') 'The future development direction should not be framed as another isolated engineering tool.'
AddDoc (New-Paragraph 'Decision requested: approve the Phase 2 configuration core plus one bounded Continuous Naval Assurance pilot. Use the 26-week delivery WBS and objective KPI gates as the acceptance model.' 22 $true '162033' 160) 'Decision requested: approve the Phase 2 configuration core plus one bounded Continuous Naval Assurance pilot.'

AddDoc (New-Heading1 '2. Why the Future Direction Must Change') '2. Why the Future Direction Must Change'
AddDoc (New-Paragraph 'Marine development is moving beyond single-tool productivity. Shipyards need confidence that every engineering decision can be traced from requirement to design revision, BOM/effectivity, affected scenario, solver evidence, approval and operations feedback. Today, that loop is fragmented across authoring systems, PLM systems, calculation tools, handover repositories and operations systems.') 'Marine development is moving beyond single-tool productivity.'
AddDoc (New-Bullet 'Current gap: no vendor owns the complete change-to-assurance-to-operations loop.') 'Current gap: no vendor owns the complete change-to-assurance-to-operations loop.'
AddDoc (New-Bullet 'Consequence: teams rely on manual reconciliation, local knowledge and disconnected approvals.') 'Consequence: teams rely on manual reconciliation, local knowledge and disconnected approvals.'
AddDoc (New-Bullet 'Future AVEVA response: own the cross-domain decision control plane, not every specialist calculation.') 'Future AVEVA response: own the cross-domain decision control plane, not every specialist calculation.'
AddDoc (New-Bullet 'AI boundary: use AI for diagnosis, recommendations and gate holds only; humans and class authorities retain release authority.') 'AI boundary: use AI for diagnosis, recommendations and gate holds only.'

AddDoc (New-Heading1 '3. Target Product Position') '3. Target Product Position'
AddDoc (New-Paragraph 'The proposed future product category is an Open AVEVA Marine Decision Control Plane. It does not remove AVEVA E3D, Marine, Hull, Draw, NAPA or other approved tools. Instead, it makes their outputs governable and auditable across the lifecycle.') 'The proposed future product category is an Open AVEVA Marine Decision Control Plane.'
AddDoc (New-Heading2 'Configuration') 'Configuration'
AddDoc (New-Bullet 'Own BOM, effectivity, baseline, revision and affected-object relationships.') 'Own BOM, effectivity, baseline, revision and affected-object relationships.'
AddDoc (New-Bullet 'Customer value: the shipyard can understand which hull, block, option, date or operational state is affected.') 'Customer value: the shipyard can understand which hull, block, option, date or operational state is affected.'
AddDoc (New-Heading2 'Impact and Evidence') 'Impact and Evidence'
AddDoc (New-Bullet 'Link requirement, design, BOM, solver scenario, diagnostics, approval and audit history.') 'Link requirement, design, BOM, solver scenario, diagnostics, approval and audit history.'
AddDoc (New-Bullet 'Customer value: class/customer reviewers can replay the evidence behind the decision.') 'Customer value: class/customer reviewers can replay the evidence behind the decision.'
AddDoc (New-Heading2 'Approval and Operations Learning') 'Approval and Operations Learning'
AddDoc (New-Bullet 'Control evaluate, review, hold, promote and override states with role-gated authority.') 'Control evaluate, review, hold, promote and override states with role-gated authority.'
AddDoc (New-Bullet 'Feed AIM, PI, CONNECT and operations calibration back into future engineering decisions.') 'Feed AIM, PI, CONNECT and operations calibration back into future engineering decisions.'

AddDoc (New-Heading1 '4. Architecture Direction') '4. Architecture Direction'
AddDoc (New-Paragraph 'The architecture must be explained carefully because it is a major feasibility point for development teams. The correct boundary is hybrid.') 'The architecture must be explained carefully because it is a major feasibility point for development teams.'
AddDoc (New-Bullet 'Windows Authoring Tier: AVEVA E3D, Marine, Hull and Draw remain Windows-centered authoring environments.') 'Windows Authoring Tier: AVEVA E3D, Marine, Hull and Draw remain Windows-centered authoring environments.'
AddDoc (New-Bullet 'Local Agent / Plugin Bridge: authoring tools connect through Local Agent, Plugin or PML.NET bridge patterns.') 'Local Agent / Plugin Bridge: authoring tools connect through Local Agent, Plugin or PML.NET bridge patterns.'
AddDoc (New-Bullet 'Linux Control Plane: API, SQL, AI Gate, evidence services, integration gateway and dashboards run on Linux infrastructure.') 'Linux Control Plane: API, SQL, AI Gate, evidence services, integration gateway and dashboards run on Linux infrastructure.'
AddDoc (New-Bullet 'Cloud/CONNECT Integration: AVEVA CONNECT, AIM, PI and related services are integrated through governed adapters.') 'Cloud/CONNECT Integration: AVEVA CONNECT, AIM, PI and related services are integrated through governed adapters.'
AddDoc (New-Bullet 'Certified Solver Federation: NAPA and other approved solvers remain solver-of-record systems; AVEVA records and governs the evidence.') 'Certified Solver Federation: NAPA and other approved solvers remain solver-of-record systems.'
AddDoc (New-Callout 'Implementation boundary to protect: do not assume E3D/Marine/Hull is hosted directly on Linux. Linux hosts the control plane. Windows authoring tools connect through agents, plugins or callback receivers.') 'Implementation boundary to protect: do not assume E3D/Marine/Hull is hosted directly on Linux.'

AddDoc (New-Heading1 '5. AI and Automation Governance') '5. AI and Automation Governance'
AddDoc (New-Paragraph 'AI should be positioned as a controlled assistant, not as a certification authority. This makes the proposal more credible to engineers, planners, class reviewers and security stakeholders.') 'AI should be positioned as a controlled assistant, not as a certification authority.'
AddDoc (New-Bullet 'Rule-based AI Gate: immediate deterministic validation of metadata, drawing standards, VCRM coverage, missing evidence and approval readiness.') 'Rule-based AI Gate: immediate deterministic validation.'
AddDoc (New-Bullet 'On-prem RAG: semantic search over drawings, specifications, change history, evidence packs and asset context using local data.') 'On-prem RAG: semantic search over drawings, specifications, change history, evidence packs and asset context.'
AddDoc (New-Bullet 'AVEVA CONNECT Industrial AI: secure Q&A and decision support over connected engineering and asset data when governance is ready.') 'AVEVA CONNECT Industrial AI: secure Q&A and decision support.'
AddDoc (New-Bullet 'Unified Engineering / Design AI: routing suggestions, clash alternatives and simulation acceleration, always followed by engineer and class validation.') 'Unified Engineering / Design AI: routing suggestions, clash alternatives and simulation acceleration.'
AddDoc (New-Callout 'Control principle: AI diagnoses and recommends; rule guardrails hold gates; certified solvers calculate; humans and class authorities release.') 'Control principle: AI diagnoses and recommends; humans and class authorities release.'

AddDoc (New-Heading1 '6. Continuous Naval Assurance Pilot') '6. Continuous Naval Assurance Pilot'
AddDoc (New-Paragraph 'The pilot should be deliberately narrow. It should prove one closed, auditable loop rather than attempt enterprise-wide PLM replacement.') 'The pilot should be deliberately narrow.'
AddDoc (New-Bullet 'E3D/Hull event: a design change or revision event is emitted through a Local Agent or Plugin.') 'E3D/Hull event: a design change or revision event is emitted.'
AddDoc (New-Bullet 'Configuration impact: the control plane resolves BOM, effectivity, baseline and affected scenarios.') 'Configuration impact: the control plane resolves BOM, effectivity, baseline and affected scenarios.'
AddDoc (New-Bullet 'Solver evidence: NAPA or another approved solver runs against the defined scenario.') 'Solver evidence: NAPA or another approved solver runs against the defined scenario.'
AddDoc (New-Bullet 'AI/rule gate: rules and AI assistance rank issues, missing evidence and review concerns.') 'AI/rule gate: rules and AI assistance rank issues.'
AddDoc (New-Bullet 'Approval and promote: engineer/class authority approves or holds the state.') 'Approval and promote: engineer/class authority approves or holds the state.'
AddDoc (New-Bullet 'Operations feedback: relevant operational context is linked back to future decisions.') 'Operations feedback: relevant operational context is linked back to future decisions.'

AddDoc (New-Heading1 '7. What Development Teams Must Hear') '7. What Development Teams Must Hear'
AddDoc (New-Heading2 'Build') 'Build'
AddDoc (New-Bullet 'OpenAPI contracts for objects, BOM, effectivity, baseline, change impact and assurance evidence.') 'OpenAPI contracts for objects, BOM, effectivity, baseline, change impact and assurance evidence.'
AddDoc (New-Bullet 'Local Agent or Plugin callbacks from Windows authoring tools.') 'Local Agent or Plugin callbacks from Windows authoring tools.'
AddDoc (New-Bullet 'AI Gate provider interface: rules first, RAG next, CONNECT later.') 'AI Gate provider interface: rules first, RAG next, CONNECT later.'
AddDoc (New-Heading2 'Validate') 'Validate'
AddDoc (New-Bullet 'Seed truth cases for hull, block, compartment, weight, loading case and affected scenarios.') 'Seed truth cases for hull, block, compartment, weight, loading case and affected scenarios.'
AddDoc (New-Bullet 'E2E tests for evaluate -> review -> promote and rollback.') 'E2E tests for evaluate -> review -> promote and rollback.'
AddDoc (New-Bullet 'KPI gates for contract coverage, accuracy, provenance, response time and security.') 'KPI gates for contract coverage, accuracy, provenance, response time and security.'
AddDoc (New-Heading2 'Avoid') 'Avoid'
AddDoc (New-Bullet 'No direct Linux hosting assumption for E3D/Marine/Hull.') 'No direct Linux hosting assumption for E3D/Marine/Hull.'
AddDoc (New-Bullet 'No big-bang replacement of certified solvers.') 'No big-bang replacement of certified solvers.'
AddDoc (New-Bullet 'No automatic AI release authority.') 'No automatic AI release authority.'
AddDoc (New-Bullet 'No unbounded process customization without approval gates.') 'No unbounded process customization without approval gates.'
AddDoc (New-Callout 'Development success statement: a small number of interfaces become reliable enough that planning, engineering and class reviewers can trust the decision state.') 'Development success statement: interfaces become reliable enough that reviewers can trust the decision state.'

AddDoc (New-Heading1 '8. What Planning Teams Must Hear') '8. What Planning Teams Must Hear'
AddDoc (New-Bullet 'Product position: AVEVA owns the control plane above specialist tools. The customer keeps tool choice, but AVEVA owns lifecycle decision, evidence and approval orchestration.') 'Product position: AVEVA owns the control plane above specialist tools.'
AddDoc (New-Bullet 'Customer value: fewer blind handovers, faster change impact review, visible assurance evidence, explainable approvals and replayable history from design to operations.') 'Customer value: fewer blind handovers, faster change impact review, visible assurance evidence.'
AddDoc (New-Bullet 'Roadmap discipline: a 26-week gate model prevents wishful scope. Each capability must pass measurable acceptance before promotion.') 'Roadmap discipline: a 26-week gate model prevents wishful scope.'
AddDoc (New-Bullet 'Decision request: approve the bounded Phase 2 + Continuous Naval Assurance increment and agree owners, KPI thresholds, assumption boundaries and pilot scenario.') 'Decision request: approve the bounded Phase 2 + Continuous Naval Assurance increment.'
AddDoc (New-Callout 'Planning success statement: the future product direction becomes easy to explain: AVEVA Marine becomes the trusted system for seeing, approving and learning from every major shipbuilding change.') 'Planning success statement: AVEVA Marine becomes the trusted system for major shipbuilding change.'

AddDoc (New-Heading1 '9. Delivery Roadmap') '9. Delivery Roadmap'
AddDoc (New-Bullet '0-2 weeks: readiness, backlog, source data, roles and test environment.') '0-2 weeks: readiness, backlog, source data, roles and test environment.'
AddDoc (New-Bullet '2-6 weeks: Phase 1 hardening, OIDC, dashboard, CI, seed truth and current gates.') '2-6 weeks: Phase 1 hardening.'
AddDoc (New-Bullet '6-12 weeks: configuration core, BOM, effectivity, baseline and APIs.') '6-12 weeks: configuration core.'
AddDoc (New-Bullet '12-16 weeks: change impact, ECR/ECO/MOC, impact graph and promote flow.') '12-16 weeks: change impact.'
AddDoc (New-Bullet '16-22 weeks: naval assurance, E3D/Hull event, solver adapters and evidence pack.') '16-22 weeks: naval assurance.'
AddDoc (New-Bullet '22-26 weeks: UAT, adoption, class/customer review and pilot closeout.') '22-26 weeks: UAT, adoption, class/customer review and pilot closeout.'

AddDoc (New-Heading1 '10. KPI Acceptance Model') '10. KPI Acceptance Model'
AddDoc (New-Bullet 'API contract coverage: 100%. Every committed capability is visible, testable and integrable.') 'API contract coverage: 100%.'
AddDoc (New-Bullet 'E2E gate pass rate: >=95%. Promote/rollback behavior is reliable under CI.') 'E2E gate pass rate: >=95%.'
AddDoc (New-Bullet 'BOM/effectivity accuracy: >=99%. Configuration decisions are trusted.') 'BOM/effectivity accuracy: >=99%.'
AddDoc (New-Bullet 'Impact graph recall: >=95%. Affected objects and naval scenarios are not missed.') 'Impact graph recall: >=95%.'
AddDoc (New-Bullet 'Solver-run provenance: 100%. Evidence is replayable and defensible.') 'Solver-run provenance: 100%.'
AddDoc (New-Bullet 'Evidence-pack completeness: 100%. Approval has the required rule trace and affected-object context.') 'Evidence-pack completeness: 100%.'
AddDoc (New-Bullet 'AI diagnostic precision: >=85%. AI explanations are useful without becoming release authority.') 'AI diagnostic precision: >=85%.'
AddDoc (New-Bullet 'Dashboard/approval response: <=5 seconds. Reviewers can use the system in real approval workflows.') 'Dashboard/approval response: <=5 seconds.'

AddDoc (New-Heading1 '11. Risks and Mitigations') '11. Risks and Mitigations'
AddDoc (New-Bullet 'Scope creep: pilot one E3D/Hull scenario and federate solvers instead of rewriting them.') 'Scope creep: pilot one E3D/Hull scenario.'
AddDoc (New-Bullet 'Solver or class acceptance: keep certified solver as calculation truth; preserve provenance and human/class release.') 'Solver or class acceptance: keep certified solver as calculation truth.'
AddDoc (New-Bullet 'Data/model mapping complexity: use seeded truth cases for hull, compartment, weight, loading case and rule links.') 'Data/model mapping complexity: use seeded truth cases.'
AddDoc (New-Bullet 'AI overreach or security concern: limit AI to recommendation and gate-hold support; enforce OIDC/RBAC, signed evidence and audit trail.') 'AI overreach or security concern: limit AI to recommendation and gate-hold support.'
AddDoc (New-Bullet 'Adoption and ownership: use a cross-functional review board with product planning, naval architecture, solution architecture and development.') 'Adoption and ownership: use a cross-functional review board.'

AddDoc (New-Heading1 '12. Meeting Facilitation Guide') '12. Meeting Facilitation Guide'
AddDoc (New-Bullet 'Slides 3-6: establish current gaps and public product signals.') 'Slides 3-6: establish current gaps and public product signals.'
AddDoc (New-Bullet 'Slides 7-15: explain the winning product direction, including AI and lightweight routing.') 'Slides 7-15: explain the winning product direction.'
AddDoc (New-Bullet 'Slides 16-29: translate strategy into architecture, pilot scope, WBS and objective acceptance criteria.') 'Slides 16-29: translate strategy into architecture, pilot scope, WBS and objective acceptance criteria.'
AddDoc (New-Bullet 'Slides 30-34: re-score the future state, address ownership/risk/evidence and ask for the decision.') 'Slides 30-34: re-score the future state and ask for the decision.'
AddDoc (New-Bullet 'Slides 38-41: use when the room needs a more explicit persuasion structure for development or planning stakeholders.') 'Slides 38-41: use when the room needs a more explicit persuasion structure.'

AddDoc (New-Heading1 '13. Recommended Closing Argument') '13. Recommended Closing Argument'
AddDoc (New-Callout 'The next differentiator for AVEVA Marine is not another isolated tool feature. It is a governed, auditable, operations-aware control plane that lets customers see, explain, approve and replay every major engineering state transition.') 'The next differentiator for AVEVA Marine is not another isolated tool feature.'
AddDoc (New-Paragraph 'Final ask: align on the product direction, approve the bounded pilot scope, assign owners, validate the first customer/class evidence scenario and use KPI gates to decide promotion.' 22 $true '162033' 160) 'Final ask: align on the product direction, approve the bounded pilot scope, assign owners, validate the first customer/class evidence scenario and use KPI gates to decide promotion.'
AddDoc (New-Paragraph 'Reference deck: Future_Industrial_PLM_Meeting_Deck_EN_V7.pptx. This Word document is written as standalone English meeting material for development and planning teams.' 18 $false '666666' 120) 'Reference deck: Future_Industrial_PLM_Meeting_Deck_EN_V7.pptx.'

$sectPr = '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1080" w:right="1080" w:bottom="1080" w:left="1080" w:header="720" w:footer="720" w:gutter="0"/><w:cols w:space="720"/><w:docGrid w:linePitch="360"/></w:sectPr>'
$documentXml = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" mc:Ignorable="w14 w15 wp14">
<w:body>
$($parts -join "`r`n")
$sectPr
</w:body>
</w:document>
"@

$contentTypes = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>
'@

$rels = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>
'@

$docRels = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>
'@

$core = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
<dc:title>AVEVA Marine Development Future Direction</dc:title>
<dc:subject>Development and Planning Meeting Materials</dc:subject>
<dc:creator>Codex</dc:creator>
<cp:keywords>AVEVA Marine; PLM; development; planning; control plane; naval assurance</cp:keywords>
<dc:description>English standalone Word meeting materials for AVEVA Marine future development direction.</dc:description>
<cp:lastModifiedBy>Codex</cp:lastModifiedBy>
<dcterms:created xsi:type="dcterms:W3CDTF">$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssZ')</dcterms:created>
<dcterms:modified xsi:type="dcterms:W3CDTF">$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssZ')</dcterms:modified>
</cp:coreProperties>
"@

$app = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
<Application>Microsoft Word</Application>
<DocSecurity>0</DocSecurity>
<ScaleCrop>false</ScaleCrop>
<Company></Company>
<LinksUpToDate>false</LinksUpToDate>
<SharedDoc>false</SharedDoc>
<HyperlinksChanged>false</HyperlinksChanged>
<AppVersion>16.0000</AppVersion>
</Properties>
'@

[System.IO.File]::WriteAllText((Join-Path $buildDir '[Content_Types].xml'), $contentTypes, [System.Text.Encoding]::UTF8)
[System.IO.File]::WriteAllText((Join-Path $buildDir '_rels\.rels'), $rels, [System.Text.Encoding]::UTF8)
[System.IO.File]::WriteAllText((Join-Path $buildDir 'word\document.xml'), $documentXml, [System.Text.Encoding]::UTF8)
[System.IO.File]::WriteAllText((Join-Path $buildDir 'word\_rels\document.xml.rels'), $docRels, [System.Text.Encoding]::UTF8)
[System.IO.File]::WriteAllText((Join-Path $buildDir 'docProps\core.xml'), $core, [System.Text.Encoding]::UTF8)
[System.IO.File]::WriteAllText((Join-Path $buildDir 'docProps\app.xml'), $app, [System.Text.Encoding]::UTF8)

Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

function Add-ZipEntry {
    param(
        [System.IO.Compression.ZipArchive]$Archive,
        [string]$EntryName,
        [string]$SourcePath
    )
    $entry = $Archive.CreateEntry($EntryName, [System.IO.Compression.CompressionLevel]::Optimal)
    $entryStream = $entry.Open()
    $fileStream = [System.IO.File]::OpenRead($SourcePath)
    try {
        $fileStream.CopyTo($entryStream)
    } finally {
        $fileStream.Dispose()
        $entryStream.Dispose()
    }
}

$zipStream = [System.IO.File]::Open($tempDocx, [System.IO.FileMode]::CreateNew)
$archive = New-Object System.IO.Compression.ZipArchive($zipStream, [System.IO.Compression.ZipArchiveMode]::Create)
try {
    Add-ZipEntry -Archive $archive -EntryName '[Content_Types].xml' -SourcePath (Join-Path $buildDir '[Content_Types].xml')
    Add-ZipEntry -Archive $archive -EntryName '_rels/.rels' -SourcePath (Join-Path $buildDir '_rels\.rels')
    Add-ZipEntry -Archive $archive -EntryName 'docProps/core.xml' -SourcePath (Join-Path $buildDir 'docProps\core.xml')
    Add-ZipEntry -Archive $archive -EntryName 'docProps/app.xml' -SourcePath (Join-Path $buildDir 'docProps\app.xml')
    Add-ZipEntry -Archive $archive -EntryName 'word/document.xml' -SourcePath (Join-Path $buildDir 'word\document.xml')
    Add-ZipEntry -Archive $archive -EntryName 'word/_rels/document.xml.rels' -SourcePath (Join-Path $buildDir 'word\_rels\document.xml.rels')
} finally {
    $archive.Dispose()
    $zipStream.Dispose()
}

New-Item -ItemType Directory -Path $backupDir | Out-Null
foreach ($p in @($finalDocx, $finalTxt)) {
    if (Test-Path -LiteralPath $p) {
        Copy-Item -LiteralPath $p -Destination (Join-Path $backupDir (Split-Path $p -Leaf)) -Force
    }
}
Copy-Item -LiteralPath $tempDocx -Destination $finalDocx -Force
[System.IO.File]::WriteAllLines($finalTxt, $plain, [System.Text.Encoding]::UTF8)

$zip = [System.IO.Compression.ZipFile]::OpenRead($tempDocx)
try {
    $entries = $zip.Entries | Select-Object -ExpandProperty FullName
} finally {
    $zip.Dispose()
}

[pscustomobject]@{
    TempDocx = $tempDocx
    FinalDocx = $finalDocx
    FinalTxt = $finalTxt
    TempLength = (Get-Item -LiteralPath $tempDocx).Length
    FinalLength = (Get-Item -LiteralPath $finalDocx).Length
    EntryCount = $entries.Count
    HasDocumentXml = ($entries -contains 'word/document.xml')
    BackupDir = $backupDir
} | Format-List
