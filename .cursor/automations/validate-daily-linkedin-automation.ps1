# Validate Daily LinkedIn Cursor automation files
# Run from PowerShell 5.1 on Windows.

$ErrorActionPreference = "Stop"

$AutomationDir = Join-Path $env:USERPROFILE ".cursor\automations\daily-linkedin-marine-plm-post"
$TomlPath = Join-Path $AutomationDir "automation.toml"
$MemoryPath = Join-Path $AutomationDir "memory.md"
$PromptPath = Join-Path $AutomationDir "prompt.md"

$requiredPromptSections = @(
    "Cursor Cloud Agent addendum",
    "Quiet daily LinkedIn workflow",
    "New AI-era developer leadership angle",
    "Beyond Vibe Coding",
    "Lead Developer Career Guide",
    "Post depth rule",
    "Professional mini-image rule",
    "Image overlap QA rule",
    "ready for final posting",
    "Vibe Coding"
)

$failures = @()

foreach ($path in @($TomlPath, $MemoryPath, $PromptPath)) {
    if (-not (Test-Path $path)) {
        $failures += "Missing file: $path"
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

if ($failures.Count -gt 0) {
    Write-Host "VALIDATION FAILED"
    $failures | ForEach-Object { Write-Host "  - $_" }
    exit 1
}

Write-Host "VALIDATION PASSED"
Write-Host "  automation.toml: $TomlPath"
Write-Host "  memory.md: $MemoryPath"
Write-Host "  prompt.md: $PromptPath"
Write-Host "  Schedule: daily 09:00 (cron 0 9 * * *)"
Write-Host "  Register at: https://cursor.com/automations/new"
exit 0
