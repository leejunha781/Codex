# Sync Daily LinkedIn Claude QA review automation to live Codex install
# Run from PowerShell 5.1 on Windows after pulling repo changes.

$ErrorActionPreference = "Stop"

$AutomationName = "daily-linkedin-claude-review"
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

New-Item -ItemType Directory -Force -Path $TargetDir | Out-Null

if (-not (Test-Path (Join-Path $SourceDir "automation.toml"))) {
    throw "Source automation.toml not found at $SourceDir"
}

Copy-IfDifferent -Source (Join-Path $SourceDir "automation.toml") -Destination (Join-Path $TargetDir "automation.toml")
if (Test-Path (Join-Path $SourceDir "memory.md")) {
    Copy-IfDifferent -Source (Join-Path $SourceDir "memory.md") -Destination (Join-Path $TargetDir "memory.md")
}

Write-Host ""
Write-Host "Claude QA review automation ready at $TargetDir"
Write-Host "Schedule: daily 09:20 | posting QA | Fail -> Blocked"
Write-Host "Verify ACTIVE in Codex Automations (codex://automations)"
Write-Host "Repo branch default: $Branch"
