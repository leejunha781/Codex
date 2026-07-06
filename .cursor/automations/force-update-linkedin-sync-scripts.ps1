# Force-overwrite LinkedIn sync scripts from GitHub memory branch.
# Use when local scripts are stale and sync-both still fails with Copy-Item self-overwrite.
#
# Run:
#   powershell -NoProfile -ExecutionPolicy Bypass -File C:\Users\namma\.cursor\automations\force-update-linkedin-sync-scripts.ps1

$ErrorActionPreference = "Stop"

$Branch = "memory"
$Repo = "leejunha781/Codex"
$CursorDir = Join-Path $env:USERPROFILE ".cursor\automations"
$CodexDir = Join-Path $env:USERPROFILE ".codex\automations"

$Downloads = @(
    @{ Url = "https://raw.githubusercontent.com/$Repo/$Branch/.cursor/automations/sync-daily-linkedin-automation.ps1"; Dest = Join-Path $CursorDir "sync-daily-linkedin-automation.ps1" },
    @{ Url = "https://raw.githubusercontent.com/$Repo/$Branch/.cursor/automations/sync-both-linkedin-automations.ps1"; Dest = Join-Path $CursorDir "sync-both-linkedin-automations.ps1" },
    @{ Url = "https://raw.githubusercontent.com/$Repo/$Branch/.cursor/automations/fetch-cursor-linkedin-scripts-from-github.ps1"; Dest = Join-Path $CursorDir "fetch-cursor-linkedin-scripts-from-github.ps1" },
    @{ Url = "https://raw.githubusercontent.com/$Repo/$Branch/.cursor/automations/force-update-linkedin-sync-scripts.ps1"; Dest = Join-Path $CursorDir "force-update-linkedin-sync-scripts.ps1" },
    @{ Url = "https://raw.githubusercontent.com/$Repo/$Branch/.codex/automations/sync-daily-linkedin-automation.ps1"; Dest = Join-Path $CodexDir "sync-daily-linkedin-automation.ps1" },
    @{ Url = "https://raw.githubusercontent.com/$Repo/$Branch/.codex/automations/sync-daily-linkedin-mirror-automation.ps1"; Dest = Join-Path $CodexDir "sync-daily-linkedin-mirror-automation.ps1" },
    @{ Url = "https://raw.githubusercontent.com/$Repo/$Branch/.codex/automations/daily-linkedin-marine-plm-post/automation.toml"; Dest = Join-Path $CodexDir "daily-linkedin-marine-plm-post\automation.toml" },
    @{ Url = "https://raw.githubusercontent.com/$Repo/$Branch/.codex/automations/daily-linkedin-marine-plm-post/memory.md"; Dest = Join-Path $CodexDir "daily-linkedin-marine-plm-post\memory.md" },
    @{ Url = "https://raw.githubusercontent.com/$Repo/$Branch/.codex/automations/daily-linkedin-mirror-and-post/automation.toml"; Dest = Join-Path $CodexDir "daily-linkedin-mirror-and-post\automation.toml" }
)

New-Item -ItemType Directory -Force -Path $CursorDir | Out-Null
New-Item -ItemType Directory -Force -Path $CodexDir | Out-Null
New-Item -ItemType Directory -Force -Path (Join-Path $CodexDir "daily-linkedin-marine-plm-post") | Out-Null
New-Item -ItemType Directory -Force -Path (Join-Path $CodexDir "daily-linkedin-mirror-and-post") | Out-Null

foreach ($item in $Downloads) {
    Invoke-WebRequest -Uri $item.Url -UseBasicParsing -OutFile $item.Dest
    Write-Host "OK  $($item.Dest)"
}

$marker = Select-String -Path (Join-Path $CursorDir "sync-daily-linkedin-automation.ps1") -Pattern "Copy-IfDifferent" -SimpleMatch
if (-not $marker) {
    throw "Update failed: sync-daily-linkedin-automation.ps1 still missing Copy-IfDifferent"
}

Write-Host ""
Write-Host "UPDATE COMPLETE — run:"
Write-Host "  cd $CursorDir"
Write-Host "  .\sync-both-linkedin-automations.ps1"
