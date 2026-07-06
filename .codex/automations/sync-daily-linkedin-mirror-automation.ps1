# Sync Daily LinkedIn mirror-and-post automation to live Codex install
# Run from PowerShell 5.1 on Windows after pulling repo changes.

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

$SourceDir = Join-Path $PSScriptRoot "daily-linkedin-mirror-and-post"
$TargetDir = Join-Path $env:USERPROFILE ".codex\automations\daily-linkedin-mirror-and-post"
$LegacyDir = Join-Path $env:USERPROFILE ".codex\automations\daily-linkedin-mirror-runs"

if (-not (Test-Path (Join-Path $SourceDir "automation.toml"))) {
    throw "Source automation.toml not found at $SourceDir"
}

New-Item -ItemType Directory -Force -Path $TargetDir | Out-Null
Copy-IfDifferent -Source (Join-Path $SourceDir "automation.toml") -Destination (Join-Path $TargetDir "automation.toml")

if (Test-Path $LegacyDir) {
    Write-Host "Note: legacy automation dir still exists at $LegacyDir — disable 'Daily LinkedIn Mirror Runs' in Codex UI"
}

Write-Host "Mirror-and-post automation ready at $TargetDir"
Write-Host "Schedule: daily 09:35 | Mirror: C:\Users\namma\Documents\Codex\YYYY-MM-DD\<topic-slug>\"
