# Sync Daily LinkedIn mirror-and-post automation to live Codex install
# Run from PowerShell 5.1 on Windows after pulling repo changes.

$ErrorActionPreference = "Stop"

$SourceDir = Join-Path $PSScriptRoot "daily-linkedin-mirror-and-post"
$TargetDir = Join-Path $env:USERPROFILE ".codex\automations\daily-linkedin-mirror-and-post"
$LegacyDir = Join-Path $env:USERPROFILE ".codex\automations\daily-linkedin-mirror-runs"

if (-not (Test-Path (Join-Path $SourceDir "automation.toml"))) {
    throw "Source automation.toml not found at $SourceDir"
}

New-Item -ItemType Directory -Force -Path $TargetDir | Out-Null
Copy-Item -Path (Join-Path $SourceDir "automation.toml") -Destination $TargetDir -Force

if (Test-Path $LegacyDir) {
    Write-Host "Note: legacy automation dir still exists at $LegacyDir — disable 'Daily LinkedIn Mirror Runs' in Codex UI"
}

Write-Host "Synced mirror-and-post automation to $TargetDir"
Write-Host "Schedule: daily 09:35 (mirror + LinkedIn Windows app auto-post)"
Write-Host "Mirror target: C:\Users\namma\Documents\Codex\YYYY-MM-DD\<topic-slug>\"
