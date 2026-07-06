# Run LinkedIn mirror + post test for today's mirrored run.
# Requires: Windows, LinkedIn app logged in.
#
# Usage:
#   powershell -NoProfile -File C:\Users\namma\.cursor\automations\test-linkedin-mirror-and-post.ps1
#   powershell -NoProfile -File C:\Users\namma\.cursor\automations\test-linkedin-mirror-and-post.ps1 -Date 2026-07-06 -Topic leo-satellite-terminal-commissioning-gates

[CmdletBinding()]
param(
    [string]$Date = (Get-Date).ToString("yyyy-MM-dd"),
    [string]$Topic = "",
    [string]$Branch = "cursor/daily-linkedin-test-non-plm-0681",
    [switch]$MirrorOnly
)

$ErrorActionPreference = "Stop"
$RepoRoot = "C:\Users\namma"
$AutomationsDir = Join-Path $env:USERPROFILE ".cursor\automations"
$MirrorScript = Join-Path $AutomationsDir "mirror-linkedin-runs.ps1"

if (Test-Path (Join-Path $RepoRoot ".git")) {
    Write-Host "Fetching latest test run from origin/$Branch ..."
    git -C $RepoRoot fetch origin $Branch 2>&1 | Out-Null
    git -C $RepoRoot merge --ff-only "origin/$Branch" 2>&1 | Out-Null
    if ($LASTEXITCODE -ne 0) {
        Write-Host "WARN: fast-forward merge failed; continuing with local repo state"
    }
}

if (-not (Test-Path $MirrorScript)) {
    throw "Missing mirror script: $MirrorScript"
}

Write-Host "=== Step 1: Mirror repo runs to Documents\Codex ==="
& $MirrorScript -Pull -IncludeRemoteBranches -Date $Date

$summaryPath = Join-Path $AutomationsDir "cache\linkedin-mirror\latest-mirror-runs.json"
if (-not (Test-Path $summaryPath)) {
    throw "Mirror summary not found: $summaryPath"
}

$runs = Get-Content $summaryPath -Raw | ConvertFrom-Json
if ($runs -isnot [array]) { $runs = @($runs) }

$selected = $null
if ($Topic) {
    $selected = $runs | Where-Object { $_.topic -eq $Topic } | Select-Object -First 1
} else {
    $selected = $runs | Select-Object -Last 1
}

if (-not $selected) {
    throw "No mirrored run found for date=$Date topic=$Topic"
}

$postPath = $selected.postPath
$imagePath = $selected.imagePath
$windowsDir = $selected.windowsMirrorPath

if (-not (Test-Path $postPath)) { throw "Missing post: $postPath" }
if (-not (Test-Path $imagePath)) { throw "Missing image: $imagePath" }

Write-Host ""
Write-Host "=== Mirror verified ==="
Write-Host "Topic:       $($selected.topic)"
Write-Host "Folder:      $windowsDir"
Write-Host "Post:        $postPath"
Write-Host "Image:       $imagePath"

if ($MirrorOnly) { exit 0 }

Write-Host ""
Write-Host "=== Step 2: LinkedIn Windows app auto-post ==="
Write-Host "Open Codex -> Automations -> 'Daily LinkedIn Mirror and Post' -> Run now."
Write-Host "It will auto-click Post per post-linkedin-windows-app-prompt.md"
Write-Host "Guide: $AutomationsDir\post-linkedin-windows-app-prompt.md"
