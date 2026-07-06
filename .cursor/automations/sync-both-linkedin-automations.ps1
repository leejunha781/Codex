# Sync Daily LinkedIn automation to both Codex and Cursor installs
# Run from PowerShell 5.1 on Windows after pulling repo changes.

$ErrorActionPreference = "Stop"

$RepoRoot = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
$CodexSync = Join-Path $RepoRoot ".codex\automations\sync-daily-linkedin-automation.ps1"
$CodexMirrorSync = Join-Path $RepoRoot ".codex\automations\sync-daily-linkedin-mirror-automation.ps1"
$CursorSync = Join-Path $PSScriptRoot "sync-daily-linkedin-automation.ps1"
$MirrorScript = Join-Path $PSScriptRoot "mirror-linkedin-runs.ps1"

if (Test-Path $CodexSync) {
    & $CodexSync
} else {
    Write-Warning "Codex sync script not found: $CodexSync"
}

if (Test-Path $CodexMirrorSync) {
    & $CodexMirrorSync
} else {
    Write-Warning "Codex mirror sync script not found: $CodexMirrorSync"
}

& $CursorSync

if (Test-Path $MirrorScript) {
    Write-Host ""
    Write-Host "Running local mirror (repo runs -> Documents\Codex)..."
    & $MirrorScript -Pull -IncludeRemoteBranches
} else {
    Write-Warning "Mirror script not found: $MirrorScript"
}

Write-Host ""
Write-Host "Codex, Cursor, and mirror automation synced."
