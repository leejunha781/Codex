# Validate Daily LinkedIn Cursor automation files
# Run from PowerShell 5.1 on Windows.

$ErrorActionPreference = "Stop"

$AutomationDir = Join-Path $env:USERPROFILE ".cursor\automations\daily-linkedin-marine-plm-post"
$AutomationsRoot = Join-Path $env:USERPROFILE ".cursor\automations"
$MirrorScript = Join-Path $AutomationsRoot "mirror-linkedin-runs.ps1"
$PostPrompt = Join-Path $AutomationsRoot "post-linkedin-windows-app-prompt.md"
$SyncBothScript = Join-Path $AutomationsRoot "sync-both-linkedin-automations.ps1"
$ValidateScript = Join-Path $AutomationsRoot "validate-daily-linkedin-automation.ps1"
$SyncScript = Join-Path $AutomationsRoot "sync-daily-linkedin-automation.ps1"
$InstallScript = Join-Path $AutomationsRoot "install-cursor-linkedin-automation-scripts.ps1"
$TomlPath = Join-Path $AutomationDir "automation.toml"
$MemoryPath = Join-Path $AutomationDir "memory.md"
$PromptPath = Join-Path $AutomationDir "prompt.md"

$IntegrationGuide = Join-Path $AutomationDir "figma-notion-claude-integration.md"
$NotionConfig = Join-Path $AutomationDir "notion-config.json"
$ClaudeReviewPrompt = Join-Path $AutomationDir "claude-review-prompt.md"

$requiredPromptSections = @(
    "Cursor Cloud Agent addendum",
    "Figma + Notion + Claude integration",
    "Notion topic selection rule",
    "Figma infographic rule",
    "Claude review handoff",
    "mirror-linkedin-runs.ps1",
    "daily-linkedin-mirror-and-post",
    "daily-linkedin-claude-review",
    "post-linkedin-windows-app-prompt",
    "Quiet daily LinkedIn workflow",
    "New AI-era developer leadership angle",
    "Beyond Vibe Coding",
    "Lead Developer Career Guide",
    "Post depth rule",
    "Professional mini-image rule",
    "Photo diversity rule",
    "Image generation rule",
    "Reference-grade flat vector style rule",
    "YARD",
    "VESSEL",
    "Image overlap QA rule",
    "ready for final posting",
    "Vibe Coding"
)

$failures = @()

$RequiredPaths = @($TomlPath, $MemoryPath, $PromptPath, $MirrorScript, $PostPrompt, $SyncBothScript, $ValidateScript, $SyncScript, $IntegrationGuide, $NotionConfig, $ClaudeReviewPrompt)
$OptionalPaths = @($InstallScript)

foreach ($path in $RequiredPaths) {
    if (-not (Test-Path $path)) {
        $failures += "Missing file: $path"
    }
}

foreach ($path in $OptionalPaths) {
    if (-not (Test-Path $path)) {
        Write-Host "WARN optional file missing: $path"
    }
}

if ($failures.Count -eq 0) {
    $toml = Get-Content $TomlPath -Raw
    foreach ($section in $requiredPromptSections) {
        if ($toml -notmatch [regex]::Escape($section)) {
            $failures += "automation.toml missing required prompt section: $section"
        }
    }
    if ($toml -notmatch 'platform = "cursor"') {
        $failures += 'automation.toml missing platform = "cursor"'
    }
    if ($toml -notmatch 'status = "ACTIVE"') {
        $failures += 'automation.toml status is not ACTIVE'
    }
    if ($toml -notmatch 'repository = "leejunha781/Codex"') {
        $failures += 'automation.toml repository mismatch'
    }
}

$memory = if (Test-Path $MemoryPath) { Get-Content $MemoryPath -Raw } else { "" }
if ($memory -notmatch "Topic rotation log") {
    $failures += "memory.md missing Topic rotation log section"
}

if ($failures.Count -eq 0 -and (Test-Path $PromptPath)) {
    $prompt = Get-Content $PromptPath -Raw
    if ($prompt -notmatch [regex]::Escape("Photo diversity rule")) {
        $failures += "prompt.md missing Photo diversity rule section"
    }
}

if ($failures.Count -gt 0) {
    Write-Host "VALIDATION FAILED"
    $failures | ForEach-Object { Write-Host "  - $_" }
    exit 1
}

Write-Host "VALIDATION PASSED"
Write-Host "  automation.toml: $TomlPath"
Write-Host "  memory.md: $MemoryPath"
Write-Host "  prompt.md: $PromptPath"
Write-Host "  mirror script: $MirrorScript"
Write-Host "  post prompt: $PostPrompt"
Write-Host "  sync-both: $SyncBothScript"
Write-Host "  validate: $ValidateScript"
Write-Host "  sync: $SyncScript"
Write-Host "  install: $InstallScript"
Write-Host "  integration: $IntegrationGuide"
Write-Host "  notion config: $NotionConfig"
Write-Host "  claude review: $ClaudeReviewPrompt"
Write-Host "  Schedule: daily 09:00 cloud + 09:20 Claude QA + 09:35 local mirror and LinkedIn auto-post"
Write-Host "  Register at: https://cursor.com/automations/new"
exit 0
