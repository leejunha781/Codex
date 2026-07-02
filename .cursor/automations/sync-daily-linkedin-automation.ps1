# Sync Daily LinkedIn automation to live Cursor install
# Run from PowerShell 5.1 on Windows after pulling repo changes.

$ErrorActionPreference = "Stop"

$SourceDir = Join-Path $PSScriptRoot "daily-linkedin-marine-plm-post"
$TargetDir = Join-Path $env:USERPROFILE ".cursor\automations\daily-linkedin-marine-plm-post"

if (-not (Test-Path (Join-Path $SourceDir "automation.toml"))) {
    throw "Source automation.toml not found at $SourceDir"
}

New-Item -ItemType Directory -Force -Path $TargetDir | Out-Null
New-Item -ItemType Directory -Force -Path (Join-Path $SourceDir "runs") | Out-Null

Copy-Item -Path (Join-Path $SourceDir "automation.toml") -Destination $TargetDir -Force
Copy-Item -Path (Join-Path $SourceDir "memory.md") -Destination $TargetDir -Force
Copy-Item -Path (Join-Path $SourceDir "prompt.md") -Destination $TargetDir -Force
Copy-Item -Path (Join-Path $SourceDir "cursor-cloud-registration.md") -Destination $TargetDir -Force

Write-Host "Synced Cursor automation files to $TargetDir"
Write-Host ""
Write-Host "Next steps in Cursor:"
Write-Host "  1. Open https://cursor.com/automations/new"
Write-Host "  2. Create scheduled automation: daily 09:00 (cron 0 9 * * *)"
Write-Host "  3. Repository: leejunha781/Codex (branch master)"
Write-Host "  4. Enable Memories and Computer use"
Write-Host "  5. Paste prompt from $TargetDir\prompt.md"
Write-Host "  6. Save and activate"
Write-Host ""
Write-Host "See cursor-cloud-registration.md for full checklist."
