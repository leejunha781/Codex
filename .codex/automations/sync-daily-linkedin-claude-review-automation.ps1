# Sync Daily LinkedIn Claude QA review automation to live Codex install
# Run from PowerShell 5.1 on Windows after pulling repo changes.

$ErrorActionPreference = "Stop"

$AutomationName = "daily-linkedin-claude-review"
$Repo = "leejunha781/Codex"
$Branch = "cursor/linkedin-figma-notion-claude-0681"

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

function Ensure-AutomationToml {
    param(
        [string]$TargetDir,
        [string]$RelativeRepoPath
    )

    New-Item -ItemType Directory -Force -Path $TargetDir | Out-Null
    $tomlPath = Join-Path $TargetDir "automation.toml"
    if (Test-Path $tomlPath) {
        return $tomlPath
    }

    $url = "https://raw.githubusercontent.com/$Repo/$Branch/$RelativeRepoPath/automation.toml"
    Write-Host "FETCH $url"
    Invoke-WebRequest -Uri $url -UseBasicParsing -OutFile $tomlPath
    return $tomlPath
}

function Copy-IfDifferent {
    param([string]$Source, [string]$Destination)
    $destParent = Split-Path $Destination -Parent
    if ($destParent) { New-Item -ItemType Directory -Force -Path $destParent | Out-Null }
    if (Test-SamePath $Source $Destination) {
        Write-Host "SKIP (repo = live install): $Destination"
        return
    }
    Copy-Item -Path $Source -Destination $Destination -Force
    Write-Host "OK   $Destination"
}

$SourceDir = Join-Path $PSScriptRoot $AutomationName
$TargetDir = Join-Path $env:USERPROFILE ".codex\automations\$AutomationName"
$RelativeRepoPath = ".codex/automations/$AutomationName"

if (-not (Test-Path (Join-Path $SourceDir "automation.toml"))) {
    Write-Host "Source automation.toml not in repo path; ensuring live install from GitHub..."
    Ensure-AutomationToml -TargetDir $TargetDir -RelativeRepoPath $RelativeRepoPath | Out-Null
} else {
    Copy-IfDifferent -Source (Join-Path $SourceDir "automation.toml") -Destination (Join-Path $TargetDir "automation.toml")
}

if (-not (Test-Path (Join-Path $TargetDir "automation.toml"))) {
    throw "automation.toml still missing at $TargetDir"
}

Write-Host ""
Write-Host "Claude QA review automation ready at $TargetDir"
Write-Host "Schedule: daily 09:20 | commissioning-gates QA | Fail -> Blocked"
Write-Host "Verify ACTIVE in Codex Automations (codex://automations)"
