# Force-overwrite LinkedIn automation scripts from GitHub (bypasses repo=live SKIP).
# Use when sync-both shows SKIP and scripts stay stale.
#
# Run:
#   powershell -NoProfile -ExecutionPolicy Bypass -File "$env:USERPROFILE\.cursor\automations\force-update-linkedin-sync-scripts.ps1"

$ErrorActionPreference = "Stop"

$Branch = "memory"
$Repo = "leejunha781/Codex"
$CursorBase = "https://raw.githubusercontent.com/$Repo/$Branch/.cursor/automations"
$CodexBase = "https://raw.githubusercontent.com/$Repo/$Branch/.codex/automations"
$CursorDir = Join-Path $env:USERPROFILE ".cursor\automations"
$CodexDir = Join-Path $env:USERPROFILE ".codex\automations"

function Download-GitHubFile {
    param([string]$Url, [string]$Dest)
    $parent = Split-Path $Dest -Parent
    if ($parent) { New-Item -ItemType Directory -Force -Path $parent | Out-Null }
    Invoke-WebRequest -Uri $Url -UseBasicParsing -OutFile $Dest -Headers @{ "User-Agent" = "Codex-LinkedIn-Automation" }
    Write-Host "OK  $Dest"
}

$Downloads = @(
    @{ Url = "$CursorBase/sync-daily-linkedin-automation.ps1"; Dest = Join-Path $CursorDir "sync-daily-linkedin-automation.ps1" },
    @{ Url = "$CursorBase/sync-both-linkedin-automations.ps1"; Dest = Join-Path $CursorDir "sync-both-linkedin-automations.ps1" },
    @{ Url = "$CursorBase/validate-daily-linkedin-automation.ps1"; Dest = Join-Path $CursorDir "validate-daily-linkedin-automation.ps1" },
    @{ Url = "$CursorBase/mirror-linkedin-runs.ps1"; Dest = Join-Path $CursorDir "mirror-linkedin-runs.ps1" },
    @{ Url = "$CursorBase/strip-linkedin-image-c2pa.ps1"; Dest = Join-Path $CursorDir "strip-linkedin-image-c2pa.ps1" },
    @{ Url = "$CursorBase/fetch-cursor-linkedin-scripts-from-github.ps1"; Dest = Join-Path $CursorDir "fetch-cursor-linkedin-scripts-from-github.ps1" },
    @{ Url = "$CursorBase/run-linkedin-automation-setup.ps1"; Dest = Join-Path $CursorDir "run-linkedin-automation-setup.ps1" },
    @{ Url = "$CursorBase/force-update-linkedin-sync-scripts.ps1"; Dest = Join-Path $CursorDir "force-update-linkedin-sync-scripts.ps1" },
    @{ Url = "$CursorBase/post-linkedin-windows-app-prompt.md"; Dest = Join-Path $CursorDir "post-linkedin-windows-app-prompt.md" },
    @{ Url = "$CursorBase/daily-linkedin-marine-plm-post/automation.toml"; Dest = Join-Path $CursorDir "daily-linkedin-marine-plm-post\automation.toml" },
    @{ Url = "$CursorBase/daily-linkedin-marine-plm-post/claude-review-prompt.md"; Dest = Join-Path $CursorDir "daily-linkedin-marine-plm-post\claude-review-prompt.md" },
    @{ Url = "$CursorBase/daily-linkedin-marine-plm-post/linkedin-reference-style-commissioning-gates.md"; Dest = Join-Path $CursorDir "daily-linkedin-marine-plm-post\linkedin-reference-style-commissioning-gates.md" },
    @{ Url = "$CursorBase/daily-linkedin-marine-plm-post/linear-config.json"; Dest = Join-Path $CursorDir "daily-linkedin-marine-plm-post\linear-config.json" },
    @{ Url = "$CursorBase/daily-linkedin-marine-plm-post/notion-config.json"; Dest = Join-Path $CursorDir "daily-linkedin-marine-plm-post\notion-config.json" },
    @{ Url = "$CodexBase/sync-daily-linkedin-automation.ps1"; Dest = Join-Path $CodexDir "sync-daily-linkedin-automation.ps1" },
    @{ Url = "$CodexBase/sync-daily-linkedin-mirror-automation.ps1"; Dest = Join-Path $CodexDir "sync-daily-linkedin-mirror-automation.ps1" },
    @{ Url = "$CodexBase/sync-daily-linkedin-claude-review-automation.ps1"; Dest = Join-Path $CodexDir "sync-daily-linkedin-claude-review-automation.ps1" },
    @{ Url = "$CodexBase/daily-linkedin-claude-review/automation.toml"; Dest = Join-Path $CodexDir "daily-linkedin-claude-review\automation.toml" },
    @{ Url = "$CodexBase/daily-linkedin-marine-plm-post/automation.toml"; Dest = Join-Path $CodexDir "daily-linkedin-marine-plm-post\automation.toml" },
    @{ Url = "$CodexBase/daily-linkedin-mirror-and-post/automation.toml"; Dest = Join-Path $CodexDir "daily-linkedin-mirror-and-post\automation.toml" }
)

New-Item -ItemType Directory -Force -Path $CursorDir | Out-Null
New-Item -ItemType Directory -Force -Path $CodexDir | Out-Null

foreach ($item in $Downloads) {
    Download-GitHubFile -Url $item.Url -Dest $item.Dest
}

$mirrorMarker = Select-String -Path (Join-Path $CursorDir "mirror-linkedin-runs.ps1") -Pattern "cmd.exe /c" -SimpleMatch
if (-not $mirrorMarker) {
    throw "Update failed: mirror-linkedin-runs.ps1 still missing cmd.exe git wrapper"
}

$syncMarker = Select-String -Path (Join-Path $CursorDir "sync-both-linkedin-automations.ps1") -Pattern "-IncludeRemoteBranches" -SimpleMatch
if (-not $syncMarker) {
    throw "Update failed: sync-both-linkedin-automations.ps1 still stale"
}

Write-Host ""
Write-Host "FORCE UPDATE COMPLETE"
Write-Host "Next:"
Write-Host "  powershell -NoProfile -ExecutionPolicy Bypass -File `"$CursorDir\validate-daily-linkedin-automation.ps1`""
