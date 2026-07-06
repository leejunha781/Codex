# Sync Daily LinkedIn automation to live Cursor install
# Run from PowerShell 5.1 on Windows after pulling repo changes.

$ErrorActionPreference = "Stop"

$SourceDir = Join-Path $PSScriptRoot "daily-linkedin-marine-plm-post"
$TargetDir = Join-Path $env:USERPROFILE ".cursor\automations\daily-linkedin-marine-plm-post"
$InstallDir = Join-Path $env:USERPROFILE ".cursor\automations"

if (-not (Test-Path (Join-Path $SourceDir "automation.toml"))) {
    throw "Source automation.toml not found at $SourceDir"
}

New-Item -ItemType Directory -Force -Path $TargetDir | Out-Null
New-Item -ItemType Directory -Force -Path (Join-Path $SourceDir "runs") | Out-Null
New-Item -ItemType Directory -Force -Path $InstallDir | Out-Null

Copy-Item -Path (Join-Path $SourceDir "automation.toml") -Destination $TargetDir -Force
Copy-Item -Path (Join-Path $SourceDir "memory.md") -Destination $TargetDir -Force
Copy-Item -Path (Join-Path $SourceDir "prompt.md") -Destination $TargetDir -Force
Copy-Item -Path (Join-Path $SourceDir "cursor-cloud-registration.md") -Destination $TargetDir -Force

& (Join-Path $PSScriptRoot "install-linkedin-mirror.ps1")

Write-Host ""
Write-Host "Synced Cursor automation files to $TargetDir"
Write-Host "Installed mirror tools to $InstallDir"
Write-Host ""
Write-Host "Mirror latest LinkedIn run (from anywhere):"
Write-Host "  powershell -ExecutionPolicy Bypass -File `"$InstallDir\mirror-linkedin-run-to-codex.ps1`" -Latest"
Write-Host ""
Write-Host "See WINDOWS-MIRROR-GUIDE.md for troubleshooting."
