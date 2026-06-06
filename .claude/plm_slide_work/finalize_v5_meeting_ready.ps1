$ErrorActionPreference = 'Stop'

$proposal = 'D:\이력서\AVEVA - Marine Principal Technical Support & Consultant – PLM SME, Busan\Proposal'
$v5Path = Join-Path $proposal 'Future_Industrial_PLM_Meeting_Deck_EN_V5.pptx'
$defaultPath = Join-Path $proposal 'Future_Industrial_PLM_Meeting_Deck_EN.pptx'
$finalPath = Join-Path $proposal 'Future_Industrial_PLM_Meeting_Deck_EN_Final.pptx'
$backupDir = Join-Path $proposal '_backup_20260606_final_meeting_ready_v5'

New-Item -ItemType Directory -Path $backupDir -Force | Out-Null
$timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
Copy-Item -LiteralPath $v5Path -Destination (Join-Path $backupDir ("Future_Industrial_PLM_Meeting_Deck_EN_V5_BEFORE_FINAL_MEETING_READY_$timestamp.pptx")) -Force
if (Test-Path -LiteralPath $defaultPath) {
    Copy-Item -LiteralPath $defaultPath -Destination (Join-Path $backupDir ("Future_Industrial_PLM_Meeting_Deck_EN_DEFAULT_BEFORE_FINAL_MEETING_READY_$timestamp.pptx")) -Force
}

function Get-ShapeByName {
    param($Slide, [string]$Name)
    foreach ($shape in @($Slide.Shapes)) {
        if ($shape.Name -eq $Name) { return $shape }
    }
    throw "Shape not found: slide $($Slide.SlideIndex), $Name"
}

function Set-ShapeText {
    param($Slide, [string]$Name, [string]$Text)
    $shape = Get-ShapeByName -Slide $Slide -Name $Name
    if (-not ($shape.HasTextFrame)) { throw "Shape has no text frame: slide $($Slide.SlideIndex), $Name" }
    $shape.TextFrame.TextRange.Text = $Text
}

$pp = $null
$pres = $null

try {
    $pp = New-Object -ComObject PowerPoint.Application
    $pp.Visible = -1
    $pres = $pp.Presentations.Open($v5Path, $false, $false, $false)

    # Page 2: agenda must explicitly account for the new AI strategy and routing pilot pages.
    $s = $pres.Slides.Item(2)
    Set-ShapeText $s 'TextBox 30' 'Meeting thesis, win-above patterns, AI strategy and deterministic routing pilot'
    Set-ShapeText $s 'TextBox 50' 'Approve configuration core, AI gate/routing pilot and assurance pilot'

    # Page 15: frame the solver stack as a bounded pilot candidate, not a locked platform mandate.
    $s = $pres.Slides.Item(15)
    Set-ShapeText $s 'TextBox 29' 'Candidate solver stack: scipy + networkx + 3D voxel grid; useful routing suggestions without full Generative Design AI in the first pilot.'

    # Page 29: KPI model must measure the new routing pilot as well as generic AI diagnostics.
    $s = $pres.Slides.Item(29)
    Set-ShapeText $s 'TextBox 7' 'Objective gates measure correctness, evidence provenance, route feasibility and decision speed - not subjective progress'
    Set-ShapeText $s 'TextBox 96' 'Solver / routing provenance'
    Set-ShapeText $s 'TextBox 100' 'Version, input, route candidate, solver result and approval retained'
    Set-ShapeText $s 'TextBox 110' 'AI / routing diagnostic precision'
    Set-ShapeText $s 'TextBox 114' 'Useful explanations and route suggestions without release authority'

    # Page 31: review ownership must include planning/development responsibility for AI/routing scope.
    $s = $pres.Slides.Item(31)
    Set-ShapeText $s 'TextBox 14' 'Roadmap · MVP scope · AI/routing pilot outcomes · KPI thresholds'
    Set-ShapeText $s 'TextBox 53' 'Implementation · AI gate/routing service · automation · security · performance'
    Set-ShapeText $s 'TextBox 56' '1 KPI + evidence dashboard 2 Failing configuration / assurance gates 3 AI/routing pilot issues 4 Scope-change decisions 5 Promote completed capabilities + update risk burndown'

    # Page 32: risk language must keep AI and generated routes advisory.
    $s = $pres.Slides.Item(32)
    Set-ShapeText $s 'TextBox 109' 'AI / routing overreach + security'
    Set-ShapeText $s 'TextBox 110' 'AI explains only; route candidates remain advisory; OIDC/RBAC, signed evidence and full audit trail.'
    Set-ShapeText $s 'TextBox 115' 'Naval architect, engineering, IT and planning jointly own AI/routing pilot and KPI decisions.'

    # Page 33: assumption boundary now includes routing-service and provider-choice validation.
    $s = $pres.Slides.Item(33)
    Set-ShapeText $s 'TextBox 78' 'Deployment boundary: E3D/Marine/Unified Engineering authoring is Windows-centered; Linux hosts API, SQL, AI Gate, evidence, routing service and integration. CONNECT, AIM, PI, Edge and solver choices vary by product/config; validate scope with vendors, IT and class authorities.'

    # Page 34: final decision must unlock the AI Gate + route-proposal pilot introduced on pages 13-15.
    $s = $pres.Slides.Item(34)
    Set-ShapeText $s 'TextBox 8' 'Approve the bounded next increment that proves Windows Authoring + Linux Control Plane, AI Gate and routing pilot in a real shipbuilding change scenario'
    Set-ShapeText $s 'TextBox 14' 'Phase 2 configuration core + bounded AI/routing + Continuous Naval Assurance pilot'
    Set-ShapeText $s 'TextBox 22' 'One E3D/Hull Local Agent change-to-evidence + route proposal pilot slice'
    Set-ShapeText $s 'TextBox 31' 'Harden reference + AI Gate'
    Set-ShapeText $s 'TextBox 32' 'OIDC · seed truth · route rules · E2E gates'
    Set-ShapeText $s 'Rounded Rectangle 50' 'Affected scenarios + route proposal'
    Set-ShapeText $s 'Rounded Rectangle 63' 'WHY NOW | The executable reference is ready; the next differentiator is controlled configuration + AI/routing evidence + lifecycle learning.'

    $pres.Save()
}
finally {
    if ($pres) { $pres.Close() }
    if ($pp) { $pp.Quit() }
}

Copy-Item -LiteralPath $v5Path -Destination $defaultPath -Force
Copy-Item -LiteralPath $v5Path -Destination $finalPath -Force

Get-Item -LiteralPath $v5Path, $defaultPath, $finalPath |
    Select-Object Name, Length, LastWriteTime
