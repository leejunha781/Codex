# Mirror a daily LinkedIn automation run from the repo to the canonical Codex output folder.
# Run on Windows (Cursor Pro / Codex local) after pulling cloud-run artifacts.
#
# Usage (from repo root):
#   powershell -ExecutionPolicy Bypass -File .\.cursor\automations\mirror-linkedin-run-to-codex.ps1 -Latest
#   powershell -ExecutionPolicy Bypass -File .\.cursor\automations\mirror-linkedin-run-to-codex.ps1 -Date 2026-07-06 -TopicSlug marine-plm-design-change-control-governance
#   powershell -ExecutionPolicy Bypass -File .\.cursor\automations\mirror-linkedin-run-to-codex.ps1 -Verify

param(
    [string]$Date,
    [string]$TopicSlug,
    [switch]$Latest,
    [switch]$Verify,
    [string]$RepoRoot,
    [string]$CodexRoot
)

$ErrorActionPreference = "Stop"

function Resolve-RepoRoot {
    param([string]$StartDir)

    $knownRepoCandidates = @(
        "C:\Users\namma\Documents\Codex",
        (Join-Path $env:USERPROFILE "Documents\Codex"),
        (Join-Path $env:USERPROFILE "OneDrive\Documents\Codex")
    )

    $configPath = Join-Path $env:USERPROFILE ".cursor\automations\codex-repo.path"
    if (Test-Path -LiteralPath $configPath) {
        $configured = (Get-Content -LiteralPath $configPath -Raw).Trim()
        if ($configured) { $knownRepoCandidates = @($configured) + $knownRepoCandidates }
    }

    foreach ($candidate in ($knownRepoCandidates | Select-Object -Unique)) {
        $runsCandidate = Join-Path $candidate ".cursor\automations\daily-linkedin-marine-plm-post\runs"
        if (Test-Path -LiteralPath $runsCandidate) {
            return (Resolve-Path -LiteralPath $candidate).Path
        }
    }

    $dir = (Resolve-Path -LiteralPath $StartDir).Path
    for ($i = 0; $i -lt 8; $i++) {
        $runsCandidate = Join-Path $dir ".cursor\automations\daily-linkedin-marine-plm-post\runs"
        if (Test-Path -LiteralPath $runsCandidate) {
            return $dir
        }
        $parent = Split-Path -Path $dir -Parent
        if (-not $parent -or $parent -eq $dir) { break }
        $dir = $parent
    }

    throw @"
Could not locate Codex repo runs folder.

Searched from: $StartDir
Expected runs at: <repo>\.cursor\automations\daily-linkedin-marine-plm-post\runs

Fix:
  1) Clone/pull repo to C:\Users\namma\Documents\Codex
  2) git fetch origin
  3) git checkout cursor/daily-linkedin-marine-plm-04c5
  4) git pull
  5) Run install once:
       powershell -ExecutionPolicy Bypass -File C:\Users\namma\Documents\Codex\.cursor\automations\install-linkedin-mirror.ps1
  6) Then from anywhere:
       powershell -ExecutionPolicy Bypass -File $env:USERPROFILE\.cursor\automations\mirror-linkedin-run-to-codex.ps1 -Latest
"@
}

function Resolve-CodexRoot {
    param([string]$Preferred)

    if ($Preferred -and (Test-Path -LiteralPath (Split-Path $Preferred -Parent))) {
        return $Preferred
    }

    $candidates = @(
        "C:\Users\namma\Documents\Codex",
        (Join-Path $env:USERPROFILE "Documents\Codex"),
        (Join-Path $env:USERPROFILE "OneDrive\Documents\Codex")
    ) | Select-Object -Unique

    foreach ($candidate in $candidates) {
        $parent = Split-Path $candidate -Parent
        if (Test-Path -LiteralPath $parent) {
            return $candidate
        }
    }

    return "C:\Users\namma\Documents\Codex"
}

if (-not $RepoRoot) {
    $RepoRoot = Resolve-RepoRoot -StartDir $PSScriptRoot
}
if (-not $CodexRoot) {
    $CodexRoot = Resolve-CodexRoot -Preferred "C:\Users\namma\Documents\Codex"
}

$RunsRoot = Join-Path $RepoRoot ".cursor\automations\daily-linkedin-marine-plm-post\runs"

Write-Host "Mirror diagnostics:"
Write-Host "  RepoRoot:   $RepoRoot"
Write-Host "  RunsRoot:   $RunsRoot"
Write-Host "  CodexRoot:  $CodexRoot"
Write-Host ""

if (-not (Test-Path -LiteralPath $RunsRoot)) {
    throw "Runs directory not found: $RunsRoot"
}

function Get-LatestRun {
    $dateDirs = @(Get-ChildItem -Path $RunsRoot -Directory | Sort-Object Name -Descending)
    if ($dateDirs.Count -eq 0) {
        throw "No date folders under $RunsRoot. Pull the latest branch (cursor/daily-linkedin-marine-plm-04c5 or merged main)."
    }

    foreach ($dateDir in $dateDirs) {
        $topicDirs = @(Get-ChildItem -Path $dateDir.FullName -Directory | Sort-Object LastWriteTime -Descending)
        if ($topicDirs.Count -gt 0) {
            return [PSCustomObject]@{
                Date = $dateDir.Name
                TopicSlug = $topicDirs[0].Name
                SourceDir = $topicDirs[0].FullName
            }
        }
    }

    throw "No topic folders found under $RunsRoot"
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

Write-Host "Selected run:"
Write-Host "  Date:       $Date"
Write-Host "  Topic:      $TopicSlug"
Write-Host "  SourceDir:  $SourceDir"
Write-Host ""

if (-not (Test-Path -LiteralPath $SourceDir)) {
    throw "Source run not found: $SourceDir`nPull branch cursor/daily-linkedin-marine-plm-04c5 or merge PR #16, then retry."
}

$postSource = Join-Path $SourceDir "linkedin-post.md"
$imageCandidates = @(Get-ChildItem -Path $SourceDir -Filter "*-infographic.png" -File -ErrorAction SilentlyContinue)
$manifestSource = Join-Path $SourceDir "ready-for-posting.json"

if (-not (Test-Path -LiteralPath $postSource)) {
    throw "Post markdown not found: $postSource"
}
if ($imageCandidates.Count -eq 0) {
    throw @"
Infographic PNG not found in: $SourceDir

The repo .gitignore ignores *.png by default. Ensure the branch includes the force-tracked infographic, then:
  git fetch origin
  git checkout cursor/daily-linkedin-marine-plm-04c5
  git pull
  git ls-files .cursor/automations/daily-linkedin-marine-plm-post/runs/$Date/$TopicSlug/
"@
}

$imageSource = $imageCandidates[0].FullName
$DestDir = Join-Path (Join-Path $CodexRoot $Date) $TopicSlug

Write-Host "Destination:"
Write-Host "  DestDir:    $DestDir"
Write-Host "  Post:       $(Join-Path $DestDir 'linkedin-post.md')"
Write-Host "  Image:      $(Join-Path $DestDir "$TopicSlug-infographic.png")"
Write-Host ""

if ($Verify) {
    Write-Host "VERIFY ONLY — no files copied."
    exit 0
}

New-Item -ItemType Directory -Force -Path $DestDir | Out-Null

$postDest = Join-Path $DestDir "linkedin-post.md"
$imageDest = Join-Path $DestDir "$TopicSlug-infographic.png"
$manifestDest = Join-Path $DestDir "ready-for-posting.json"

Copy-Item -LiteralPath $postSource -Destination $postDest -Force
Copy-Item -LiteralPath $imageSource -Destination $imageDest -Force
if (Test-Path -LiteralPath $manifestSource) {
    Copy-Item -LiteralPath $manifestSource -Destination $manifestDest -Force
} else {
    $manifest = @{
        topic = $TopicSlug
        date = $Date
        status = "ready_for_final_posting"
        repoPath = $SourceDir
        windowsPath = $DestDir
        postPath = $postDest
        imagePath = $imageDest
        mirroredAt = (Get-Date).ToString("o")
        linkedInApp = "Windows LinkedIn app only; do not click Post without user confirmation"
    }
    $manifest | ConvertTo-Json -Depth 4 | Set-Content -Path $manifestDest -Encoding UTF8
}

Write-Host "SUCCESS — mirrored LinkedIn run:"
Write-Host "  Post:       $postDest"
Write-Host "  Image:      $imageDest"
Write-Host "  Manifest:   $manifestDest"
Write-Host ""
Write-Host "Next: open LinkedIn Windows app -> Start a post -> paste post text -> attach image from above path."
