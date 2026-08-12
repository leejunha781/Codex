# One-shot LinkedIn automation setup (bypasses Restricted execution policy)
# Run from any PowerShell window:
#   powershell -NoProfile -ExecutionPolicy Bypass -File "$env:USERPROFILE\.cursor\automations\run-linkedin-automation-setup.ps1"

$ErrorActionPreference = "Stop"

$AutomationsDir = Join-Path $env:USERPROFILE ".cursor\automations"
$SyncBoth = Join-Path $AutomationsDir "sync-both-linkedin-automations.ps1"
$Validate = Join-Path $AutomationsDir "validate-daily-linkedin-automation.ps1"
$Fetch = Join-Path $AutomationsDir "fetch-cursor-linkedin-scripts-from-github.ps1"

if (-not (Test-Path $AutomationsDir)) {
    New-Item -ItemType Directory -Force -Path $AutomationsDir | Out-Null
}

if (-not (Test-Path $SyncBoth)) {
    Write-Host "Scripts missing — fetching from GitHub..."
    if (-not (Test-Path $Fetch)) {
        $url = "https://raw.githubusercontent.com/leejunha781/Codex/memory/.cursor/automations/fetch-cursor-linkedin-scripts-from-github.ps1"
        Invoke-WebRequest -Uri $url -UseBasicParsing -OutFile $Fetch
    }
    & powershell -NoProfile -ExecutionPolicy Bypass -File $Fetch
}

Write-Host ""
Write-Host "=== Sync Codex + Cursor automations ==="
& powershell -NoProfile -ExecutionPolicy Bypass -File $SyncBoth

Write-Host ""
Write-Host "=== Validate ==="
& powershell -NoProfile -ExecutionPolicy Bypass -File $Validate

Write-Host ""
Write-Host "SETUP COMPLETE"
Write-Host "Next: open Codex Automations and confirm ACTIVE:"
Write-Host "  - daily-linkedin-claude-review (09:20)"
Write-Host "  - daily-linkedin-mirror-and-post (09:35)"
