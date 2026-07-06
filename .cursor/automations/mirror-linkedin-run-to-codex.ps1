# Mirror a daily LinkedIn automation run from the repo to the canonical Codex output folder.
# Run on Windows (Cursor Pro / Codex local) after pulling cloud-run artifacts.
#
# Usage:
#   .\mirror-linkedin-run-to-codex.ps1
#   .\mirror-linkedin-run-to-codex.ps1 -Date 2026-07-06 -TopicSlug marine-plm-design-change-control-governance
#   .\mirror-linkedin-run-to-codex.ps1 -Latest

param(
    [string]$Date,
    [string]$TopicSlug,
    [switch]$Latest,
    [string]$RepoRoot = $(Split-Path (Split-Path $PSScriptRoot -Parent) -Parent),
    [string]$CodexRoot = "C:\Users\namma\Documents\Codex"
)

$ErrorActionPreference = "Stop"

$RunsRoot = Join-Path $RepoRoot ".cursor\automations\daily-linkedin-marine-plm-post\runs"

if (-not (Test-Path $RunsRoot)) {
    throw "Runs directory not found: $RunsRoot"
}

function Get-LatestRun {
    $dateDirs = Get-ChildItem -Path $RunsRoot -Directory | Sort-Object Name -Descending
    foreach ($dateDir in $dateDirs) {
        $topicDirs = Get-ChildItem -Path $dateDir.FullName -Directory | Sort-Object LastWriteTime -Descending
        if ($topicDirs.Count -gt 0) {
            return [PSCustomObject]@{
                Date = $dateDir.Name
                TopicSlug = $topicDirs[0].Name
                SourceDir = $topicDirs[0].FullName
            }
        }
    }
    throw "No run folders found under $RunsRoot"
}

if ($Latest -or (-not $Date -and -not $TopicSlug)) {
    $run = Get-LatestRun
    $Date = $run.Date
    $TopicSlug = $run.TopicSlug
    $SourceDir = $run.SourceDir
} else {
    if (-not $Date -or -not $TopicSlug) {
        throw "Provide both -Date and -TopicSlug, or use -Latest."
    }
    $SourceDir = Join-Path (Join-Path $RunsRoot $Date) $TopicSlug
}

if (-not (Test-Path $SourceDir)) {
    throw "Source run not found: $SourceDir"
}

$postSource = Join-Path $SourceDir "linkedin-post.md"
$imageCandidates = Get-ChildItem -Path $SourceDir -Filter "*-infographic.png" -File
if (-not (Test-Path $postSource)) {
    throw "Post markdown not found: $postSource"
}
if ($imageCandidates.Count -eq 0) {
    throw "Infographic PNG not found in: $SourceDir"
}
$imageSource = $imageCandidates[0].FullName

$DestDir = Join-Path (Join-Path $CodexRoot $Date) $TopicSlug
New-Item -ItemType Directory -Force -Path $DestDir | Out-Null

$postDest = Join-Path $DestDir "linkedin-post.md"
$imageDest = Join-Path $DestDir "$TopicSlug-infographic.png"

Copy-Item -Path $postSource -Destination $postDest -Force
Copy-Item -Path $imageSource -Destination $imageDest -Force

$manifest = @{
    topic = $TopicSlug
    date = $Date
    status = "ready_for_final_posting"
    repoPath = $SourceDir.Replace('\', '/')
    windowsPath = $DestDir
    postPath = $postDest
    imagePath = $imageDest
    mirroredAt = (Get-Date).ToString("o")
    linkedInApp = "Windows LinkedIn app only; do not click Post without user confirmation"
}
$manifestPath = Join-Path $DestDir "ready-for-posting.json"
$manifest | ConvertTo-Json -Depth 4 | Set-Content -Path $manifestPath -Encoding UTF8

Write-Host "Mirrored LinkedIn run to Codex output folder:"
Write-Host "  Date:       $Date"
Write-Host "  Topic:      $TopicSlug"
Write-Host "  Post:       $postDest"
Write-Host "  Image:      $imageDest"
Write-Host "  Manifest:   $manifestPath"
Write-Host ""
Write-Host "Next: open LinkedIn Windows app, Start a post, paste from linkedin-post.md, attach image from above path."
