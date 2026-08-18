# Download LinkedIn automation scripts directly from GitHub (no git branch required).
# Use when git pull on main/memory does not include .cursor/automations files.
#
# Run in PowerShell 5.1:
#   powershell -NoProfile -ExecutionPolicy Bypass -Command "iwr -UseBasicParsing 'https://raw.githubusercontent.com/leejunha781/Codex/cursor/fix-linkedin-local-mirror-0681/.cursor/automations/fetch-cursor-linkedin-scripts-from-github.ps1' -OutFile $env:TEMP\fetch-linkedin.ps1; powershell -NoProfile -ExecutionPolicy Bypass -File $env:TEMP\fetch-linkedin.ps1"

[CmdletBinding()]
param(
    [string]$Branch = "memory",
    [string]$Repo = "leejunha781/Codex",
    [string]$TargetDir = (Join-Path $env:USERPROFILE ".cursor\automations")
)

$ErrorActionPreference = "Stop"

$CursorBaseRaw = "https://raw.githubusercontent.com/$Repo/$Branch/.cursor/automations"
$CodexBaseRaw = "https://raw.githubusercontent.com/$Repo/$Branch/.codex/automations"
$CursorTarget = $TargetDir
$CodexTarget = Join-Path $env:USERPROFILE ".codex\automations"

$FileMap = @(
    @{ Base = $CursorBaseRaw; Root = $CursorTarget; Rel = "sync-both-linkedin-automations.ps1" },
    @{ Base = $CursorBaseRaw; Root = $CursorTarget; Rel = "sync-daily-linkedin-automation.ps1" },
    @{ Base = $CursorBaseRaw; Root = $CursorTarget; Rel = "validate-daily-linkedin-automation.ps1" },
    @{ Base = $CursorBaseRaw; Root = $CursorTarget; Rel = "mirror-linkedin-runs.ps1" },
    @{ Base = $CursorBaseRaw; Root = $CursorTarget; Rel = "post-linkedin-windows-app-prompt.md" },
    @{ Base = $CursorBaseRaw; Root = $CursorTarget; Rel = "install-cursor-linkedin-automation-scripts.ps1" },
    @{ Base = $CursorBaseRaw; Root = $CursorTarget; Rel = "fetch-cursor-linkedin-scripts-from-github.ps1" },
    @{ Base = $CursorBaseRaw; Root = $CursorTarget; Rel = "daily-linkedin-marine-plm-post/automation.toml" },
    @{ Base = $CursorBaseRaw; Root = $CursorTarget; Rel = "daily-linkedin-marine-plm-post/memory.md" },
    @{ Base = $CursorBaseRaw; Root = $CursorTarget; Rel = "daily-linkedin-marine-plm-post/prompt.md" },
    @{ Base = $CursorBaseRaw; Root = $CursorTarget; Rel = "daily-linkedin-marine-plm-post/cursor-cloud-registration.md" },
    @{ Base = $CursorBaseRaw; Root = $CursorTarget; Rel = "daily-linkedin-marine-plm-post/claude-review-prompt.md" },
    @{ Base = $CursorBaseRaw; Root = $CursorTarget; Rel = "daily-linkedin-marine-plm-post/embedded-arm-topic-reference.md" },
    @{ Base = $CodexBaseRaw; Root = $CodexTarget; Rel = "sync-daily-linkedin-automation.ps1" },
    @{ Base = $CodexBaseRaw; Root = $CodexTarget; Rel = "sync-daily-linkedin-mirror-automation.ps1" },
    @{ Base = $CodexBaseRaw; Root = $CodexTarget; Rel = "sync-daily-linkedin-claude-review-automation.ps1" },
    @{ Base = $CodexBaseRaw; Root = $CodexTarget; Rel = "daily-linkedin-claude-review/automation.toml" },
    @{ Base = $CodexBaseRaw; Root = $CodexTarget; Rel = "daily-linkedin-claude-review/memory.md" },
    @{ Base = $CodexBaseRaw; Root = $CodexTarget; Rel = "daily-linkedin-marine-plm-post/automation.toml" },
    @{ Base = $CodexBaseRaw; Root = $CodexTarget; Rel = "daily-linkedin-marine-plm-post/memory.md" },
    @{ Base = $CodexBaseRaw; Root = $CodexTarget; Rel = "daily-linkedin-mirror-and-post/automation.toml" }
)

New-Item -ItemType Directory -Force -Path $CursorTarget | Out-Null
New-Item -ItemType Directory -Force -Path $CodexTarget | Out-Null
New-Item -ItemType Directory -Force -Path (Join-Path $CursorTarget "cache\linkedin-mirror") | Out-Null

$ok = 0
$failed = @()

foreach ($entry in $FileMap) {
    $rel = $entry.Rel
    $url = "$($entry.Base)/$($rel -replace '\\','/')"
    $dest = Join-Path $entry.Root ($rel -replace '/', [System.IO.Path]::DirectorySeparatorChar)
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
    Write-Host "FETCH COMPLETE: $ok files"
    Write-Host "  Cursor: $CursorTarget"
    Write-Host "  Codex:  $CodexTarget"
    Write-Host ""
    Write-Host "Next:"
    Write-Host "  cd $TargetDir"
    Write-Host "  .\validate-daily-linkedin-automation.ps1"
    Write-Host "  .\sync-both-linkedin-automations.ps1"
    exit 0
}

Write-Host "FETCH INCOMPLETE: $($failed.Count) failed"
exit 1
