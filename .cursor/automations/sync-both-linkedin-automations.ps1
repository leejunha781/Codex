# Sync Daily LinkedIn automation to both Codex and Cursor installs
# Run from PowerShell 5.1 on Windows after pulling repo changes.

$ErrorActionPreference = "Stop"

function Get-RepoRootFromScript {
    param([string]$ScriptDir)

    $git = Get-Command git -ErrorAction SilentlyContinue
    if ($git) {
        $root = (& git -C $ScriptDir rev-parse --show-toplevel 2>$null)
        if ($LASTEXITCODE -eq 0 -and -not [string]::IsNullOrWhiteSpace($root)) {
            return $root.Trim()
        }
    }

    return (Split-Path (Split-Path $ScriptDir -Parent) -Parent)
}

$RepoRoot = Get-RepoRootFromScript -ScriptDir $PSScriptRoot
$CodexSync = Join-Path $RepoRoot ".codex\automations\sync-daily-linkedin-automation.ps1"
$CodexMirrorSync = Join-Path $RepoRoot ".codex\automations\sync-daily-linkedin-mirror-automation.ps1"
$CodexClaudeSync = Join-Path $RepoRoot ".codex\automations\sync-daily-linkedin-claude-review-automation.ps1"
$CursorSync = Join-Path $PSScriptRoot "sync-daily-linkedin-automation.ps1"
$MirrorScript = Join-Path $PSScriptRoot "mirror-linkedin-runs.ps1"
$FetchScript = Join-Path $PSScriptRoot "fetch-cursor-linkedin-scripts-from-github.ps1"

Write-Host "Repo root: $RepoRoot"

if (Test-Path $CodexSync) {
    & $CodexSync
} else {
    Write-Warning "Codex sync script not found: $CodexSync"
    Write-Warning "Run: git pull origin memory  OR  fetch-cursor-linkedin-scripts-from-github.ps1"
}

if (Test-Path $CodexMirrorSync) {
    & $CodexMirrorSync
} else {
    Write-Warning "Codex mirror sync script not found: $CodexMirrorSync"
    if (Test-Path $FetchScript) {
        Write-Host "Attempting GitHub fetch for missing Codex automation files..."
        & $FetchScript
    }
}

if (Test-Path $CodexClaudeSync) {
    & $CodexClaudeSync
} else {
    Write-Warning "Codex Claude QA sync script not found: $CodexClaudeSync"
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
