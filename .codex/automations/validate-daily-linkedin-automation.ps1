# Validate Daily LinkedIn automation files before enabling schedule
# Run from PowerShell 5.1 on Windows.

$ErrorActionPreference = "Stop"

$AutomationDir = Join-Path $env:USERPROFILE ".codex\automations\daily-linkedin-marine-plm-post"
$TomlPath = Join-Path $AutomationDir "automation.toml"
$MemoryPath = Join-Path $AutomationDir "memory.md"

$requiredPromptSections = @(
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

if (-not (Test-Path $TomlPath)) {
    $failures += "Missing automation.toml at $TomlPath"
}
if (-not (Test-Path $MemoryPath)) {
    $failures += "Missing memory.md at $MemoryPath"
}

if ($failures.Count -eq 0) {
    $toml = Get-Content $TomlPath -Raw
    foreach ($section in $requiredPromptSections) {
        if ($toml -notmatch [regex]::Escape($section)) {
            $failures += "automation.toml missing required prompt section: $section"
        }
    }
    if ($toml -notmatch 'status = "ACTIVE"') {
        $failures += 'automation.toml status is not ACTIVE'
    }
    if ($toml -notmatch 'id = "daily-linkedin-marine-plm-post"') {
        $failures += 'automation.toml id mismatch'
    }
}

$memory = if (Test-Path $MemoryPath) { Get-Content $MemoryPath -Raw } else { "" }
if ($memory -notmatch "Topic rotation log") {
    $failures += "memory.md missing Topic rotation log section"
}
if ($memory -notmatch "Image QA") {
    $failures += "memory.md missing Image QA template"
}

if ($failures.Count -gt 0) {
    Write-Host "VALIDATION FAILED"
    $failures | ForEach-Object { Write-Host "  - $_" }
    exit 1
}

Write-Host "VALIDATION PASSED"
Write-Host "  automation.toml: $TomlPath"
Write-Host "  memory.md: $MemoryPath"
Write-Host "  Schedule: daily 09:00 (RRULE weekly all days)"
Write-Host "  Sandbox required: full access (LinkedIn app + Documents\Codex writes)"
exit 0
