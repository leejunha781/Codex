# Sync Daily LinkedIn mirror automation to live Codex install
# Run from PowerShell 5.1 on Windows after pulling repo changes.

$ErrorActionPreference = "Stop"

$SourceDir = Join-Path $PSScriptRoot "daily-linkedin-mirror-runs"
$TargetDir = Join-Path $env:USERPROFILE ".codex\automations\daily-linkedin-mirror-runs"

if (-not (Test-Path (Join-Path $SourceDir "automation.toml"))) {
    throw "Source automation.toml not found at $SourceDir"
}

New-Item -ItemType Directory -Force -Path $TargetDir | Out-Null
Copy-Item -Path (Join-Path $SourceDir "automation.toml") -Destination $TargetDir -Force

Write-Host "Synced mirror automation to $TargetDir"
Write-Host "Schedule: daily 09:30 (30 min after cloud LinkedIn generation)"
