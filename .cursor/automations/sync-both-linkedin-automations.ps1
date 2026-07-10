# Sync Daily LinkedIn automation to both Codex and Cursor installs
# Run from PowerShell 5.1 on Windows after pulling repo changes.

$ErrorActionPreference = "Stop"

$RepoRoot = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
$CodexSync = Join-Path $RepoRoot ".codex\automations\sync-daily-linkedin-automation.ps1"
$CursorSync = Join-Path $PSScriptRoot "sync-daily-linkedin-automation.ps1"

if (Test-Path $CodexSync) {
    & $CodexSync
} else {
    Write-Warning "Codex sync script not found: $CodexSync"
}

& $CursorSync
Write-Host ""
Write-Host "Both Codex and Cursor automation dirs synced."
