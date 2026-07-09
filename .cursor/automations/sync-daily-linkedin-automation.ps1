# Sync Daily LinkedIn automation to live Cursor install
# Run from PowerShell 5.1 on Windows after pulling repo changes.
# When repo root is C:\Users\namma, repo and live install paths are the same — copies are skipped safely.

$ErrorActionPreference = "Stop"

function Normalize-PathString {
    param([string]$Path)
    if ([string]::IsNullOrWhiteSpace($Path)) { return "" }
    try {
        if (Test-Path -LiteralPath $Path) {
            return (Resolve-Path -LiteralPath $Path).ProviderPath.TrimEnd('\')
        }
    } catch {}
    return $Path.TrimEnd('\')
}

function Test-SamePath {
    param([string]$Left, [string]$Right)
    return (Normalize-PathString $Left) -ieq (Normalize-PathString $Right)
}

function Copy-IfDifferent {
    param(
        [string]$Source,
        [string]$Destination
    )

    $destParent = Split-Path $Destination -Parent
    if ($destParent) {
        New-Item -ItemType Directory -Force -Path $destParent | Out-Null
    }

    if (Test-SamePath $Source $Destination) {
        Write-Host "SKIP (repo = live install): $Destination"
        return
    }

    Copy-Item -Path $Source -Destination $Destination -Force
    Write-Host "OK   $Destination"
}

$SourceDir = Join-Path $PSScriptRoot "daily-linkedin-marine-plm-post"
$TargetDir = Join-Path $env:USERPROFILE ".cursor\automations\daily-linkedin-marine-plm-post"
$CursorAutomationsDir = Join-Path $env:USERPROFILE ".cursor\automations"

if (-not (Test-Path (Join-Path $SourceDir "automation.toml"))) {
    throw "Source automation.toml not found at $SourceDir"
}

$RequiredUtilities = @(
    "sync-both-linkedin-automations.ps1",
    "sync-daily-linkedin-automation.ps1",
    "validate-daily-linkedin-automation.ps1",
    "mirror-linkedin-runs.ps1",
    "post-linkedin-windows-app-prompt.md"
)

$OptionalUtilities = @(
    "install-cursor-linkedin-automation-scripts.ps1",
    "fetch-cursor-linkedin-scripts-from-github.ps1",
    "force-update-linkedin-sync-scripts.ps1"
)

New-Item -ItemType Directory -Force -Path $TargetDir | Out-Null
New-Item -ItemType Directory -Force -Path $CursorAutomationsDir | Out-Null
New-Item -ItemType Directory -Force -Path (Join-Path $SourceDir "runs") | Out-Null

Copy-IfDifferent -Source (Join-Path $SourceDir "automation.toml") -Destination (Join-Path $TargetDir "automation.toml")
Copy-IfDifferent -Source (Join-Path $SourceDir "memory.md") -Destination (Join-Path $TargetDir "memory.md")
Copy-IfDifferent -Source (Join-Path $SourceDir "prompt.md") -Destination (Join-Path $TargetDir "prompt.md")
Copy-IfDifferent -Source (Join-Path $SourceDir "cursor-cloud-registration.md") -Destination (Join-Path $TargetDir "cursor-cloud-registration.md")
Copy-IfDifferent -Source (Join-Path $SourceDir "figma-notion-claude-integration.md") -Destination (Join-Path $TargetDir "figma-notion-claude-integration.md")
Copy-IfDifferent -Source (Join-Path $SourceDir "notion-config.json") -Destination (Join-Path $TargetDir "notion-config.json")
Copy-IfDifferent -Source (Join-Path $SourceDir "claude-review-prompt.md") -Destination (Join-Path $TargetDir "claude-review-prompt.md")
Copy-IfDifferent -Source (Join-Path $SourceDir "freelance-topic-reference.md") -Destination (Join-Path $TargetDir "freelance-topic-reference.md")
Copy-IfDifferent -Source (Join-Path $SourceDir "professional-image-design-rule.md") -Destination (Join-Path $TargetDir "professional-image-design-rule.md")
Copy-IfDifferent -Source (Join-Path $SourceDir "linkedin-reference-style-commissioning-gates.md") -Destination (Join-Path $TargetDir "linkedin-reference-style-commissioning-gates.md")
Copy-IfDifferent -Source (Join-Path $SourceDir "linear-config.json") -Destination (Join-Path $TargetDir "linear-config.json")

foreach ($utility in $RequiredUtilities) {
    $sourcePath = Join-Path $PSScriptRoot $utility
    if (-not (Test-Path $sourcePath)) {
        throw "Required utility file not found in repo: $sourcePath"
    }
    Copy-IfDifferent -Source $sourcePath -Destination (Join-Path $CursorAutomationsDir $utility)
}

foreach ($utility in $OptionalUtilities) {
    $sourcePath = Join-Path $PSScriptRoot $utility
    if (-not (Test-Path $sourcePath)) {
        Write-Host "WARN optional utility missing (skipped): $sourcePath"
        continue
    }
    Copy-IfDifferent -Source $sourcePath -Destination (Join-Path $CursorAutomationsDir $utility)
}

if (Test-SamePath $PSScriptRoot $CursorAutomationsDir) {
    Write-Host ""
    Write-Host "Repo automations dir equals live install ($CursorAutomationsDir) — file copies skipped as expected."
}

Write-Host ""
Write-Host "Cursor automation ready at $TargetDir"
Write-Host "See cursor-cloud-registration.md for activation checklist."
