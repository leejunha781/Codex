$ErrorActionPreference = 'Stop'

$proposal = 'D:\이력서\AVEVA - Marine Principal Technical Support & Consultant – PLM SME, Busan\Proposal'
$workDir = 'C:\Users\namma\.claude\plm_slide_work'
$tempDocx = Join-Path $workDir 'future_direction_meeting_brief_en.docx'
$tempPdf = Join-Path $workDir 'future_direction_meeting_brief_en.pdf'
$briefDocx = Join-Path $proposal 'Future_Industrial_PLM_Development_Planning_Meeting_Brief_EN.docx'
$briefPdf = Join-Path $proposal 'Future_Industrial_PLM_Development_Planning_Meeting_Brief_EN.pdf'
$backupDir = Join-Path $proposal ("_backup_20260606_future_direction_brief_" + (Get-Date -Format 'yyyyMMdd_HHmmss'))

New-Item -ItemType Directory -Path $backupDir | Out-Null
foreach ($existing in @($briefDocx, $briefPdf)) {
    if (Test-Path -LiteralPath $existing) {
        Copy-Item -LiteralPath $existing -Destination (Join-Path $backupDir (Split-Path $existing -Leaf)) -Force
    }
}
foreach ($generated in @($tempDocx, $tempPdf, $briefDocx, $briefPdf)) {
    if (Test-Path -LiteralPath $generated) {
        Remove-Item -LiteralPath $generated -Force
    }
}

$brief = @'
AVEVA Marine Development Future Direction
Development and Planning Meeting Brief - English Final Materials

1. Meeting Purpose
- Align development and planning teams on the future direction: an open AVEVA Marine PLM control plane above specialist engineering, solver and operations tools.
- Approve a bounded next increment: Phase 2 configuration core plus Continuous Naval Assurance pilot.
- Use objective KPI gates to decide whether each capability is ready to promote.

2. Core Thesis
AVEVA Marine should not try to replace every specialist tool. The stronger future position is to own the governed decision loop that connects Windows-based authoring, Linux-based control-plane services, certified solver evidence, human/class approval and operations feedback.

3. Persuasion Logic for the Audience
- For development: the architecture is feasible because it uses bounded FastAPI/OpenAPI contracts, PostgreSQL evidence storage, OIDC/RBAC security, Local Agent/Plugin integration and automated tests.
- For planning: the direction is differentiated because it combines Siemens-grade configuration, Dassault-grade collaboration, NAPA-grade assurance and AVEVA lifecycle context without surrendering the decision layer.
- For risk owners: AI is deliberately limited to diagnosis, recommendation and gate-hold support. Certified solvers, engineers and class authorities retain release authority.

4. Proposed Future Direction
- Build the Open AVEVA Decision Control Plane: configuration, impact, evidence, approval and operations learning.
- Keep E3D/Marine/Hull authoring on Windows and connect it through Local Agent, Plugin or PML.NET bridge patterns.
- Host API, SQL, AI Gate, evidence and integration services on the Linux control-plane layer.
- Federate NAPA and other certified solvers as sources of calculation evidence, not as systems to replace.
- Start AI with deterministic rule gates and on-prem RAG, then scale toward CONNECT and Unified Engineering AI after governance is validated.

5. Meeting Flow
- Slides 3-6 establish current gaps and public product signals.
- Slides 7-15 explain the winning product direction, including AI and lightweight routing.
- Slides 16-27 translate the direction into architecture, workflow and pilot scope.
- Slides 28-34 convert the proposal into WBS, KPI gates, risks and a decision request.
- Slides 35-37 provide glossary support for mixed development, planning and marine-domain audiences.

6. Decisions to Secure
- Agree that AVEVA should own the decision/evidence/control plane above specialist tools.
- Approve the Phase 2 configuration core and bounded Continuous Naval Assurance pilot.
- Accept the hybrid boundary: Windows authoring tier plus Linux control plane plus Cloud/CONNECT adapters.
- Confirm that AI remains advisory and gated, never the certified release authority.
- Use the 26-week WBS and KPI model as the delivery and acceptance structure.

7. Recommended Closing Message
The next differentiator for AVEVA Marine is not another isolated tool feature. It is a governed, auditable, operations-aware control plane that lets customers see, explain, approve and replay every major engineering state transition. The reference package proves the implementation pattern; the meeting should approve the bounded pilot that proves the product direction.
'@

$word = New-Object -ComObject Word.Application
$doc = $null
try {
    $word.Visible = $false
    $doc = $word.Documents.Add()
    $doc.PageSetup.TopMargin = 54
    $doc.PageSetup.BottomMargin = 54
    $doc.PageSetup.LeftMargin = 54
    $doc.PageSetup.RightMargin = 54
    $doc.Content.Font.Name = 'Aptos'
    $doc.Content.Font.Size = 10
    $doc.Content.Text = $brief
    $doc.SaveAs2($tempDocx, 16)
    $doc.ExportAsFixedFormat($tempPdf, 17)
    $doc.Close([ref]$true)
    $doc = $null
} finally {
    if ($doc -ne $null) { $doc.Close([ref]$false) }
    if ($word.Documents.Count -eq 0) { $word.Quit() }
}

Copy-Item -LiteralPath $tempDocx -Destination $briefDocx -Force
Copy-Item -LiteralPath $tempPdf -Destination $briefPdf -Force

Get-Item -LiteralPath $briefDocx, $briefPdf | Select-Object FullName, Length, LastWriteTime
