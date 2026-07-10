# Sync Daily LinkedIn automation to live Codex install
# Run from PowerShell 5.1 on Windows after pulling repo changes.

$ErrorActionPreference = "Stop"

$SourceDir = Join-Path $PSScriptRoot "daily-linkedin-marine-plm-post"
$TargetDir = Join-Path $env:USERPROFILE ".codex\automations\daily-linkedin-marine-plm-post"

if (-not (Test-Path (Join-Path $SourceDir "automation.toml"))) {
    throw "Source automation.toml not found at $SourceDir"
}

New-Item -ItemType Directory -Force -Path $TargetDir | Out-Null

Copy-Item -Path (Join-Path $SourceDir "automation.toml") -Destination $TargetDir -Force
Copy-Item -Path (Join-Path $SourceDir "memory.md") -Destination $TargetDir -Force

Write-Host "Synced automation files to $TargetDir"
Write-Host ""
Write-Host "Next steps in Codex Desktop:"
Write-Host "  1. Open Automations (codex://automations)"
Write-Host "  2. Open 'Daily LinkedIn Marine PLM Post' (or create New Schedule with same prompt)"
Write-Host "  3. Verify: ACTIVE, daily 09:00, local execution, full-access sandbox"
Write-Host "  4. Project/CWD: C:\Users\namma\Documents\Daily Linkedin Posting"
Write-Host "  5. Enable 'Prevent sleep while running' for scheduled runs"
