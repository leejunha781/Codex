# Download LinkedIn automation scripts directly from GitHub (no git branch required).
# Use when git pull on main/memory does not include .cursor/automations files.
#
# Run in PowerShell 5.1:
#   powershell -NoProfile -ExecutionPolicy Bypass -Command "iwr -UseBasicParsing 'https://raw.githubusercontent.com/leejunha781/Codex/cursor/fix-linkedin-local-mirror-0681/.cursor/automations/fetch-cursor-linkedin-scripts-from-github.ps1' -OutFile $env:TEMP\fetch-linkedin.ps1; powershell -NoProfile -ExecutionPolicy Bypass -File $env:TEMP\fetch-linkedin.ps1"

[CmdletBinding()]
param(
    [string]$Branch = "cursor/fix-linkedin-local-mirror-0681",
    [string]$Repo = "leejunha781/Codex",
    [string]$TargetDir = (Join-Path $env:USERPROFILE ".cursor\automations")
)

$ErrorActionPreference = "Stop"

$BaseRaw = "https://raw.githubusercontent.com/$Repo/$Branch/.cursor/automations"

$Files = @(
    "sync-both-linkedin-automations.ps1",
    "sync-daily-linkedin-automation.ps1",
    "validate-daily-linkedin-automation.ps1",
    "mirror-linkedin-runs.ps1",
    "post-linkedin-windows-app-prompt.md",
    "install-cursor-linkedin-automation-scripts.ps1",
    "fetch-cursor-linkedin-scripts-from-github.ps1",
    "daily-linkedin-marine-plm-post/automation.toml",
    "daily-linkedin-marine-plm-post/memory.md",
    "daily-linkedin-marine-plm-post/prompt.md",
    "daily-linkedin-marine-plm-post/cursor-cloud-registration.md"
)

New-Item -ItemType Directory -Force -Path $TargetDir | Out-Null
New-Item -ItemType Directory -Force -Path (Join-Path $TargetDir "daily-linkedin-marine-plm-post") | Out-Null
New-Item -ItemType Directory -Force -Path (Join-Path $TargetDir "cache\linkedin-mirror") | Out-Null

$ok = 0
$failed = @()

foreach ($rel in $Files) {
    $url = "$BaseRaw/$($rel -replace '\\','/')"
    $dest = Join-Path $TargetDir ($rel -replace '/', [System.IO.Path]::DirectorySeparatorChar)
    $destParent = Split-Path $dest -Parent
    if ($destParent) {
        New-Item -ItemType Directory -Force -Path $destParent | Out-Null
    }

    try {
        Invoke-WebRequest -Uri $url -UseBasicParsing -OutFile $dest
        Write-Host "OK  $rel"
        $ok++
    } catch {
        Write-Host "FAIL $rel"
        Write-Host "     $url"
        Write-Host "     $($_.Exception.Message)"
        $failed += $rel
    }
}

Write-Host ""
if ($failed.Count -eq 0) {
    Write-Host "FETCH COMPLETE: $ok files -> $TargetDir"
    Write-Host ""
    Write-Host "Next:"
    Write-Host "  cd $TargetDir"
    Write-Host "  .\validate-daily-linkedin-automation.ps1"
    Write-Host "  .\sync-both-linkedin-automations.ps1"
    exit 0
}

Write-Host "FETCH INCOMPLETE: $($failed.Count) failed"
exit 1
