$ErrorActionPreference = 'Stop'

$proposal = 'D:\이력서\AVEVA - Marine Principal Technical Support & Consultant – PLM SME, Busan\Proposal'
$sourceDeck = Join-Path $proposal 'Future_Industrial_PLM_Meeting_Deck_EN_V6.pptx'
$defaultDeck = Join-Path $proposal 'Future_Industrial_PLM_Meeting_Deck_EN.pptx'
$finalDeck = Join-Path $proposal 'Future_Industrial_PLM_Meeting_Deck_EN_Final.pptx'
$briefDocx = Join-Path $proposal 'Future_Industrial_PLM_Development_Planning_Meeting_Brief_EN.docx'
$briefPdf = Join-Path $proposal 'Future_Industrial_PLM_Development_Planning_Meeting_Brief_EN.pdf'
$workDir = 'C:\Users\namma\.claude\plm_slide_work'
$notesDump = Join-Path $workDir 'future_direction_meeting_speaker_notes_en.txt'
$runStamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$backupDir = Join-Path $proposal ("_backup_20260606_future_direction_materials_" + $runStamp)

if (-not (Test-Path -LiteralPath $sourceDeck)) {
    throw "Source deck not found: $sourceDeck"
}
if (-not (Test-Path -LiteralPath $workDir)) {
    New-Item -ItemType Directory -Path $workDir | Out-Null
}
New-Item -ItemType Directory -Path $backupDir | Out-Null

$backupTargets = @($sourceDeck, $defaultDeck, $finalDeck, $briefDocx, $briefPdf) | Where-Object { Test-Path -LiteralPath $_ }
foreach ($target in $backupTargets) {
    Copy-Item -LiteralPath $target -Destination (Join-Path $backupDir (Split-Path $target -Leaf)) -Force
}

foreach ($generated in @($briefDocx, $briefPdf)) {
    if (Test-Path -LiteralPath $generated) {
        Remove-Item -LiteralPath $generated -Force
    }
}

$notes = @{}
$notes[1] = @'
Purpose: Open the meeting by making the decision objective explicit. The audience is not being asked to approve a vague vision; they are being asked to align on a concrete future direction, pilot scope, governance model and KPI-based acceptance plan.
Talk track: Position AVEVA Marine as the platform that can connect engineering authoring, configuration control, certified solver evidence and operations feedback. Emphasize that the reference API package proves the direction is implementable, not only conceptual.
'@
$notes[2] = @'
Purpose: Explain the meeting structure before going into detail. This avoids the common objection that architecture, AI, WBS and KPI topics are disconnected.
Talk track: The deck moves from facts to strategy, then from strategy to a bounded delivery plan. Ask development to focus on feasibility and sequencing; ask planning to focus on product position, customer value and approval gates.
'@
$notes[3] = @'
Purpose: Establish the problem: no single vendor currently owns the full loop from design change to assurance evidence to operations feedback.
Talk track: Use this slide to create urgency. The opportunity is not to replace every specialist tool, but to own the controlled decision layer that connects them.
'@
$notes[4] = @'
Purpose: Place AVEVA in the competitive landscape. The key message is that AVEVA can move from a strong engineering and operations footprint into the center of governed lifecycle decisions.
Talk track: Explain that the future opportunity sits between engineering data and live operations. That is the space where a Marine PLM control plane can be differentiated.
'@
$notes[5] = @'
Purpose: Anchor the future direction in the current executable baseline. This helps development teams see that the proposal extends existing capability rather than starting from zero.
Talk track: Phase 1 already proves integration, standard data, handover readiness and dashboard value. The next step is configuration governance so every issue, BOM and change is traceable by effectivity.
'@
$notes[6] = @'
Purpose: Show that the direction follows public product signals while keeping assumptions explicit.
Talk track: AVEVA, Siemens, Dassault, Hexagon, PTC and NAPA all point toward digital thread, AI, lifecycle and assurance themes. The proposal absorbs these signals through adapters and governance instead of assuming fixed vendor roadmaps.
'@
$notes[7] = @'
Purpose: State the central thesis clearly: win above tool silos.
Talk track: AVEVA does not need to own every calculation or every authoring feature. The stronger position is to own the cross-domain decision, evidence, approval and learning loop around specialist tools.
'@
$notes[8] = @'
Purpose: Define the missing product category: an open AVEVA decision control plane.
Talk track: Highlight the five linked capabilities: configuration, impact, evidence, approval and operations feedback. This is the future direction that makes AVEVA Marine more than a set of engineering applications.
'@
$notes[9] = @'
Purpose: Translate competitor strengths into AVEVA product moves.
Talk track: The persuasive point is not that competitors are weak. It is that AVEVA can absorb their best patterns and go beyond them by combining E3D/Marine, Unified Engineering, AIM, PI, CONNECT, AI and customer-owned APIs.
'@
$notes[10] = @'
Purpose: Explain why NAPA should be federated, not displaced.
Talk track: Keep certified naval calculations where trust already exists. AVEVA should own the assurance graph around revisions, scenarios, evidence packs, approvals and lifecycle learning.
'@
$notes[11] = @'
Purpose: Show how AVEVA can match Siemens-grade configuration rigor without becoming a closed PLM stack.
Talk track: Development should hear that open APIs, effectivity, baseline, impact graph and promote gates are the core build items. Planning should hear that this creates a differentiated, operations-aware PLM position.
'@
$notes[12] = @'
Purpose: Show how AVEVA can go beyond virtual-twin presentation into executable governance.
Talk track: The important distinction is that a visual twin alone does not approve a change. The proposed direction turns context into evidence, gate decisions and auditable release states.
'@
$notes[13] = @'
Purpose: Clarify the AI position before the audience overestimates or distrusts it.
Talk track: AI assists, recommends and explains. The control plane governs; rules hold gates; certified solvers calculate; humans and class authorities release. This makes AI useful without making it unsafe.
'@
$notes[14] = @'
Purpose: Give a realistic AI adoption path.
Talk track: Start with deterministic rules because they are fast, auditable and low-risk. Add on-prem RAG for semantic search, then scale to CONNECT and Unified Engineering AI after data, security and governance are ready.
'@
$notes[15] = @'
Purpose: Make the lightweight routing concept credible and bounded.
Talk track: Present routing as an advisory deterministic pilot, not automatic design release. The value is auditable route proposals from existing control-plane infrastructure, followed by engineer and class validation.
'@
$notes[16] = @'
Purpose: Connect all strategy slides into one target architecture.
Talk track: The future AVEVA Marine direction is hybrid by design: Windows authoring remains where E3D/Marine tools run, Linux hosts the open control plane, and Cloud/CONNECT is integrated through governed adapters.
'@
$notes[17] = @'
Purpose: Prove that the concept has an executable foundation.
Talk track: Point out that FastAPI, PostgreSQL, OIDC, Docker Compose and tests already demonstrate the API-first control pattern. The meeting decision is about approving the next bounded extension.
'@
$notes[18] = @'
Purpose: Remove architecture ambiguity.
Talk track: E3D/Marine/Hull authoring is Windows-centered. Linux does not host E3D; it hosts API, SQL, AI Gate, evidence and integration services. This boundary protects feasibility and prevents wrong implementation assumptions.
'@
$notes[19] = @'
Purpose: Show the deployment model as an operating architecture.
Talk track: Emphasize that Local Agents or plugins bridge authoring tools to the control plane. This is safer than direct tool coupling and gives development a clear integration pattern.
'@
$notes[20] = @'
Purpose: Define the pilot scope so the future direction feels achievable.
Talk track: The pilot should prove one closed loop rather than attempt full enterprise PLM at once. Planning should protect scope; development should protect adapter, schema, test and evidence quality.
'@
$notes[21] = @'
Purpose: Explain requirements governance in practical terms.
Talk track: VCRM and checklist generation connect customer requirements to evidence and approval. The admin-editable model keeps planning flexibility while preserving traceability.
'@
$notes[22] = @'
Purpose: Show the end-to-end operating flow across design, PLM and approval.
Talk track: Use this slide to make the workflow concrete. Each transition is governed, each output is recorded, and AI can hold or explain a gate but cannot release the final certified state.
'@
$notes[23] = @'
Purpose: Ground the architecture in a familiar engineering artifact: 2D production drawing output.
Talk track: Show how AVEVA Draw and AutoCAD-format deliverables can be registered into the API-built PLM so drawings, MTO, checks and approvals remain connected.
'@
$notes[24] = @'
Purpose: Explain version control for E3D and Hull design data.
Talk track: The future direction requires every design state to be identifiable, comparable and promotable. This is what lets downstream impact, evidence and approval become reliable.
'@
$notes[25] = @'
Purpose: Explain Continuous Naval Assurance as an API concept, not a standalone solver.
Talk track: The control plane records solver runs, diagnostics, evidence packs and approvals. It federates trusted calculation engines while making the decision trail visible and replayable.
'@
$notes[26] = @'
Purpose: Show that the process is customer-configurable.
Talk track: Planning can change the process sequence without asking development to rewrite the platform. Development should expose safe APIs and guardrails; planning should define which process variants are allowed.
'@
$notes[27] = @'
Purpose: Explain the gate mechanism behind the proposed platform.
Talk track: The audience should see that evaluate, review and promote are deliberate lifecycle states. This prevents accidental release and creates auditable acceptance criteria.
'@
$notes[28] = @'
Purpose: Convert the future direction into a delivery plan.
Talk track: The WBS is staged to reduce risk: harden the baseline, build configuration core, add change impact, then add naval assurance and shipyard adoption. KPI gates control movement between stages.
'@
$notes[29] = @'
Purpose: Define success in measurable terms.
Talk track: These KPIs prevent subjective progress reporting. Development is measured by contracts, tests, accuracy, provenance and response time; planning is measured by whether the product value is accepted at each gate.
'@
$notes[30] = @'
Purpose: Re-score the competitive position after the proposed capabilities are applied.
Talk track: The claim is that AVEVA can become distinctive by orchestrating the full assurance loop. Keep the wording future-state and pilot-validation based so the argument remains credible.
'@
$notes[31] = @'
Purpose: Define who owns decisions after the meeting.
Talk track: The future direction requires product planning, naval architecture, solution architecture and development to share one evidence dashboard. This prevents strategy from separating from implementation.
'@
$notes[32] = @'
Purpose: Acknowledge risks before the audience raises them.
Talk track: The proposal is persuasive because it does not hide scope, solver, data, AI or adoption risks. Each risk has a bounded mitigation tied to pilot scope and governance.
'@
$notes[33] = @'
Purpose: Separate official product signals from proposal assumptions.
Talk track: This slide protects credibility. State clearly what is based on public direction and what must be validated with vendors, class authorities and the customer environment.
'@
$notes[34] = @'
Purpose: Close with the decision request.
Talk track: Ask for approval of the bounded next increment: Phase 2 configuration core plus Continuous Naval Assurance pilot. The proof point is one closed, auditable loop from E3D/Hull event to approved evidence and operations feedback.
'@
$notes[35] = @'
Purpose: Use this appendix only when terms block understanding.
Talk track: Do not present every glossary row in the main meeting. Refer to it when PLM, BOM, effectivity, baseline, VCRM or class terms need quick clarification.
'@
$notes[36] = @'
Purpose: Clarify marine and AVEVA product terminology.
Talk track: Use this appendix to align mixed development and planning audiences. It prevents confusion between authoring tools, AI tools, solvers and PLM control-plane functions.
'@
$notes[37] = @'
Purpose: Clarify architecture, security and infrastructure terminology.
Talk track: Use this appendix when discussing implementation feasibility, identity, APIs, PostgreSQL, CI, HA/DR or security. It keeps technical questions grounded in shared definitions.
'@

function Set-SlideNotes {
    param(
        [object]$Slide,
        [string]$Text
    )

    $written = $false
    try {
        $placeholder = $Slide.NotesPage.Shapes.Placeholders(2)
        $placeholder.TextFrame.TextRange.Text = $Text
        $written = $true
    } catch {}

    if (-not $written) {
        try {
            $shape = $Slide.NotesPage.Shapes.AddTextbox(1, 36, 120, 648, 360)
            $shape.TextFrame.TextRange.Text = $Text
            $written = $true
        } catch {
            throw "Unable to write notes for slide $($Slide.SlideIndex): $($_.Exception.Message)"
        }
    }
}

function Add-WordParagraph {
    param(
        [object]$Selection,
        [string]$Text,
        [single]$Size,
        [int]$Bold,
        [int]$SpaceAfter = 6
    )

    $Selection.Font.Name = 'Aptos'
    $Selection.Font.Size = $Size
    $Selection.Font.Bold = $Bold
    $Selection.ParagraphFormat.SpaceAfter = $SpaceAfter
    $Selection.TypeText($Text)
    $Selection.TypeParagraph()
}

function Add-WordBullet {
    param(
        [object]$Selection,
        [string]$Text
    )

    Add-WordParagraph -Selection $Selection -Text ("- " + $Text) -Size ([single]10) -Bold 0 -SpaceAfter 3
}

$pp = New-Object -ComObject PowerPoint.Application
$presentation = $null
try {
    $presentation = $pp.Presentations.Open($sourceDeck, $false, $false, $false)
    if ($presentation.Slides.Count -ne 37) {
        throw "Unexpected slide count: $($presentation.Slides.Count). Expected 37."
    }

    for ($i = 1; $i -le 37; $i++) {
        Set-SlideNotes -Slide $presentation.Slides.Item($i) -Text $notes[$i]
    }

    $presentation.Save()
} finally {
    if ($presentation -ne $null) { $presentation.Close() }
    if ($pp.Presentations.Count -eq 0) { $pp.Quit() }
}

$dumpLines = New-Object System.Collections.Generic.List[string]
$dumpLines.Add('Future Industrial PLM - English meeting speaker notes')
$dumpLines.Add(('Generated: {0:yyyy-MM-dd HH:mm:ss K}' -f (Get-Date)))
$dumpLines.Add(('Deck: {0}' -f $sourceDeck))
$dumpLines.Add('')
for ($i = 1; $i -le 37; $i++) {
    $dumpLines.Add("===== SLIDE $i =====")
    $dumpLines.Add($notes[$i])
    $dumpLines.Add('')
}
[System.IO.File]::WriteAllLines($notesDump, $dumpLines, [System.Text.Encoding]::UTF8)

$word = New-Object -ComObject Word.Application
$doc = $null
try {
    $word.Visible = $false
    $doc = $word.Documents.Add()
    $doc.PageSetup.TopMargin = 54
    $doc.PageSetup.BottomMargin = 54
    $doc.PageSetup.LeftMargin = 54
    $doc.PageSetup.RightMargin = 54

    $sel = $word.Selection
    Add-WordParagraph -Selection $sel -Text 'AVEVA Marine Development Future Direction' -Size ([single]18) -Bold 1 -SpaceAfter 4
    Add-WordParagraph -Selection $sel -Text 'Development and Planning Meeting Brief - English Final Materials' -Size ([single]12) -Bold 1 -SpaceAfter 12

    Add-WordParagraph -Selection $sel -Text '1. Meeting Purpose' -Size ([single]13) -Bold 1 -SpaceAfter 4
    Add-WordBullet -Selection $sel -Text 'Align development and planning teams on the future direction: an open AVEVA Marine PLM control plane above specialist engineering, solver and operations tools.'
    Add-WordBullet -Selection $sel -Text 'Approve a bounded next increment: Phase 2 configuration core plus Continuous Naval Assurance pilot.'
    Add-WordBullet -Selection $sel -Text 'Use objective KPI gates to decide whether each capability is ready to promote.'

    Add-WordParagraph -Selection $sel -Text '2. Core Thesis' -Size ([single]13) -Bold 1 -SpaceAfter 4
    Add-WordParagraph -Selection $sel -Text 'AVEVA Marine should not try to replace every specialist tool. The stronger future position is to own the governed decision loop that connects Windows-based authoring, Linux-based control-plane services, certified solver evidence, human/class approval and operations feedback.' -Size ([single]10) -Bold 0 -SpaceAfter 8

    Add-WordParagraph -Selection $sel -Text '3. Persuasion Logic for the Audience' -Size ([single]13) -Bold 1 -SpaceAfter 4
    Add-WordBullet -Selection $sel -Text 'For development: the architecture is feasible because it uses a bounded FastAPI/OpenAPI contract, PostgreSQL evidence store, OIDC/RBAC security, Local Agent/Plugin integration and automated tests.'
    Add-WordBullet -Selection $sel -Text 'For planning: the direction is differentiated because it combines Siemens-grade configuration, Dassault-grade collaboration, NAPA-grade assurance and AVEVA lifecycle context without surrendering the decision layer.'
    Add-WordBullet -Selection $sel -Text 'For risk owners: AI is deliberately limited to diagnosis, recommendation and gate-hold support. Certified solvers, engineers and class authorities retain release authority.'

    Add-WordParagraph -Selection $sel -Text '4. Proposed Future Direction' -Size ([single]13) -Bold 1 -SpaceAfter 4
    Add-WordBullet -Selection $sel -Text 'Build the Open AVEVA Decision Control Plane: configuration, impact, evidence, approval and operations learning.'
    Add-WordBullet -Selection $sel -Text 'Keep E3D/Marine/Hull authoring on Windows and connect it through Local Agent, Plugin or PML.NET bridge patterns.'
    Add-WordBullet -Selection $sel -Text 'Host API, SQL, AI Gate, evidence and integration services on the Linux control-plane layer.'
    Add-WordBullet -Selection $sel -Text 'Federate NAPA and other certified solvers as sources of calculation evidence, not as systems to replace.'
    Add-WordBullet -Selection $sel -Text 'Start AI with deterministic rule gates and on-prem RAG, then scale toward CONNECT and Unified Engineering AI after governance is validated.'

    Add-WordParagraph -Selection $sel -Text '5. Meeting Flow' -Size ([single]13) -Bold 1 -SpaceAfter 4
    Add-WordBullet -Selection $sel -Text 'Slides 3-6 establish current gaps and public product signals.'
    Add-WordBullet -Selection $sel -Text 'Slides 7-15 explain the winning product direction, including AI and lightweight routing.'
    Add-WordBullet -Selection $sel -Text 'Slides 16-27 translate the direction into architecture, workflow and pilot scope.'
    Add-WordBullet -Selection $sel -Text 'Slides 28-34 convert the proposal into WBS, KPI gates, risks and a decision request.'
    Add-WordBullet -Selection $sel -Text 'Slides 35-37 provide glossary support for mixed development, planning and marine-domain audiences.'

    Add-WordParagraph -Selection $sel -Text '6. Decisions to Secure' -Size ([single]13) -Bold 1 -SpaceAfter 4
    Add-WordBullet -Selection $sel -Text 'Agree that AVEVA should own the decision/evidence/control plane above specialist tools.'
    Add-WordBullet -Selection $sel -Text 'Approve the Phase 2 configuration core and bounded Continuous Naval Assurance pilot.'
    Add-WordBullet -Selection $sel -Text 'Accept the hybrid boundary: Windows authoring tier plus Linux control plane plus Cloud/CONNECT adapters.'
    Add-WordBullet -Selection $sel -Text 'Confirm that AI remains advisory and gated, never the certified release authority.'
    Add-WordBullet -Selection $sel -Text 'Use the 26-week WBS and KPI model as the delivery and acceptance structure.'

    Add-WordParagraph -Selection $sel -Text '7. Recommended Closing Message' -Size ([single]13) -Bold 1 -SpaceAfter 4
    Add-WordParagraph -Selection $sel -Text 'The next differentiator for AVEVA Marine is not another isolated tool feature. It is a governed, auditable, operations-aware control plane that lets customers see, explain, approve and replay every major engineering state transition. The reference package proves the implementation pattern; the meeting should approve the bounded pilot that proves the product direction.' -Size ([single]10) -Bold 0 -SpaceAfter 8

    $doc.SaveAs2($briefDocx, 16)
    $doc.ExportAsFixedFormat($briefPdf, 17)
    $doc.Close([ref]$true)
    $doc = $null
} finally {
    if ($doc -ne $null) { $doc.Close([ref]$false) }
    if ($word.Documents.Count -eq 0) { $word.Quit() }
}

Copy-Item -LiteralPath $sourceDeck -Destination $defaultDeck -Force
Copy-Item -LiteralPath $sourceDeck -Destination $finalDeck -Force

[pscustomobject]@{
    SourceDeck = $sourceDeck
    DefaultDeck = $defaultDeck
    FinalDeck = $finalDeck
    BriefDocx = $briefDocx
    BriefPdf = $briefPdf
    SpeakerNotesDump = $notesDump
    BackupDir = $backupDir
}
