# Bootstrap install of LinkedIn automation scripts into live Cursor automations dir.
# Run from PowerShell 5.1 on Windows when files are missing from:
#   C:\Users\namma\.cursor\automations\
#
# Usage (from repo after git pull):
#   powershell -NoProfile -File <repo-root>\.cursor\automations\install-cursor-linkedin-automation-scripts.ps1
#
# Or if only this file exists in repo:
#   powershell -NoProfile -File C:\Users\namma\.cursor\automations\install-cursor-linkedin-automation-scripts.ps1

$ErrorActionPreference = "Stop"

$RepoAutomationsDir = $PSScriptRoot
$LiveAutomationsDir = Join-Path $env:USERPROFILE ".cursor\automations"
$SyncScript = Join-Path $RepoAutomationsDir "sync-daily-linkedin-automation.ps1"

if (-not (Test-Path $SyncScript)) {
    throw "sync-daily-linkedin-automation.ps1 not found at $SyncScript. Git pull the latest branch first."
}

& $SyncScript

$required = @(
    "sync-both-linkedin-automations.ps1",
    "sync-daily-linkedin-automation.ps1",
    "validate-daily-linkedin-automation.ps1",
    "mirror-linkedin-runs.ps1",
    "post-linkedin-windows-app-prompt.md"
)

$missing = @()
foreach ($file in $required) {
    $livePath = Join-Path $LiveAutomationsDir $file
    if (-not (Test-Path $livePath)) {
        $missing += $livePath
    }
}

if ($missing.Count -gt 0) {
    Write-Host "INSTALL INCOMPLETE — missing:"
    $missing | ForEach-Object { Write-Host "  $_" }
    exit 1
}

Write-Host "INSTALL COMPLETE"
Write-Host "Live automations dir: $LiveAutomationsDir"
Write-Host ""
Write-Host "Next:"
Write-Host "  cd $LiveAutomationsDir"
Write-Host "  .\validate-daily-linkedin-automation.ps1"
Write-Host "  .\sync-both-linkedin-automations.ps1"
exit 0
