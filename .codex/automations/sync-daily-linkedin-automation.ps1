# Sync Daily LinkedIn automation to live Codex install
# Run from PowerShell 5.1 on Windows after pulling repo changes.

$ErrorActionPreference = "Stop"

$AutomationName = "daily-linkedin-marine-plm-post"
$Repo = "leejunha781/Codex"
$Branch = "memory"

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

function Ensure-AutomationFile {
    param(
        [string]$TargetDir,
        [string]$RelativeRepoPath,
        [string]$FileName
    )

    New-Item -ItemType Directory -Force -Path $TargetDir | Out-Null
    $destPath = Join-Path $TargetDir $FileName
    if (Test-Path $destPath) {
        return $destPath
    }

    $url = "https://raw.githubusercontent.com/$Repo/$Branch/$RelativeRepoPath/$FileName"
    Write-Host "FETCH $url"
    Invoke-WebRequest -Uri $url -UseBasicParsing -OutFile $destPath
    return $destPath
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

$sourceToml = Join-Path $SourceDir "automation.toml"
if (Test-Path $sourceToml) {
    Copy-IfDifferent -Source $sourceToml -Destination (Join-Path $TargetDir "automation.toml")
    Copy-IfDifferent -Source (Join-Path $SourceDir "memory.md") -Destination (Join-Path $TargetDir "memory.md")
} else {
    Write-Host "Source files not in repo path; ensuring live install from GitHub..."
    Ensure-AutomationFile -TargetDir $TargetDir -RelativeRepoPath $RelativeRepoPath -FileName "automation.toml" | Out-Null
    Ensure-AutomationFile -TargetDir $TargetDir -RelativeRepoPath $RelativeRepoPath -FileName "memory.md" | Out-Null
}

if (-not (Test-Path (Join-Path $TargetDir "automation.toml"))) {
    throw "automation.toml still missing at $TargetDir"
}

Write-Host ""
Write-Host "Codex automation ready at $TargetDir"
Write-Host "Next: verify ACTIVE in Codex Automations (codex://automations)"
