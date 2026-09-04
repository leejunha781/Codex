#Requires -Version 5.1
<#
.SYNOPSIS
  Fix broken Windows sandbox preflight paths and run checks.

.DESCRIPTION
  Never uses %USERPROFILE% itself as the git root (that caused
  C:\Users\namma\scripts\... and C:\Users\namma\.cursor\sandbox.json misses).

  Clones or updates github.com/leejunha781/Codex on branch
  cursor/windows-sandbox-4428 under:
    %USERPROFILE%\Codex
  or, if that folder is taken by something else:
    %USERPROFILE%\Codex-windows-sandbox

  Then runs Windows + WSL preflight scripts.

.EXAMPLE
  powershell -NoProfile -ExecutionPolicy Bypass -File Fix-Cursor-Windows-Sandbox.ps1

.EXAMPLE
  irm https://raw.githubusercontent.com/leejunha781/Codex/cursor/windows-sandbox-4428/scripts/Fix-Cursor-Windows-Sandbox.ps1 | iex
#>
[CmdletBinding()]
param(
    [string]$RepoUrl = "https://github.com/leejunha781/Codex.git",
    [string]$Branch = "cursor/windows-sandbox-4428"
)

$ErrorActionPreference = "Continue"
$fail = 0

function Write-Step([string]$Message) {
    Write-Host "[INFO] $Message" -ForegroundColor Cyan
}
function Write-Ok([string]$Message) {
    Write-Host "[PASS] $Message" -ForegroundColor Green
}
function Write-Bad([string]$Message) {
    Write-Host "[FAIL] $Message" -ForegroundColor Red
    $script:fail++
}

Write-Host ""
Write-Host "============================================================"
Write-Host " Cursor Windows Sandbox - fix paths and run preflight"
Write-Host "============================================================"
Write-Host " Never uses your user-home folder as the Codex git root."
Write-Host "============================================================"
Write-Host ""

if (-not (Get-Command git -ErrorAction SilentlyContinue)) {
    Write-Bad "git.exe not found. Install Git for Windows: https://git-scm.com/download/win"
    exit 1
}

function Test-IsCodexRepo([string]$Path) {
    $gitDir = Join-Path $Path ".git"
    if (-not (Test-Path $gitDir)) { return $false }
    Push-Location $Path
    try {
        $origin = (& git remote get-url origin 2>$null | Out-String).Trim()
        return ($origin -match "leejunha781/Codex")
    }
    finally {
        Pop-Location
    }
}

function Get-RepoDir {
    $preferred = Join-Path $env:USERPROFILE "Codex"
    $alternate = Join-Path $env:USERPROFILE "Codex-windows-sandbox"

    # Hard rule: never return $env:USERPROFILE
    if ($preferred -eq $env:USERPROFILE -or $alternate -eq $env:USERPROFILE) {
        throw "Refusing to use USERPROFILE as repo directory."
    }

    if (-not (Test-Path $preferred)) {
        return $preferred
    }

    if (Test-IsCodexRepo $preferred) {
        return $preferred
    }

    Write-Step "$preferred exists but is not leejunha781/Codex — using $alternate"
    return $alternate
}

$repoDir = Get-RepoDir
Write-Step "Repo directory: $repoDir"
Write-Step "Branch: $Branch"

# Clear stale lock only inside the chosen repo
$lock = Join-Path $repoDir ".git\index.lock"
if (Test-Path $lock) {
    Write-Step "Removing stale lock: $lock"
    Remove-Item -Force $lock -ErrorAction SilentlyContinue
}

# Also clear a stale HOME lock that blocked earlier checkouts (do not use home as repo)
$homeLock = Join-Path $env:USERPROFILE ".git\index.lock"
if (Test-Path $homeLock) {
    Write-Step "Removing stale home lock (not used as repo): $homeLock"
    Remove-Item -Force $homeLock -ErrorAction SilentlyContinue
}

if (-not (Test-IsCodexRepo $repoDir)) {
    if ((Test-Path $repoDir) -and -not (Test-Path (Join-Path $repoDir ".git"))) {
        $itemCount = @(Get-ChildItem -Force -LiteralPath $repoDir -ErrorAction SilentlyContinue).Count
        if ($itemCount -gt 0) {
            $repoDir = Join-Path $env:USERPROFILE "Codex-windows-sandbox"
            Write-Step "Non-empty non-git folder in the way — switching to $repoDir"
        }
    }

    if (Test-IsCodexRepo $repoDir) {
        Write-Step "Alternate folder already has Codex clone."
    }
    elseif (Test-Path (Join-Path $repoDir ".git")) {
        Write-Bad "$repoDir has .git but origin is not leejunha781/Codex"
        exit 1
    }
    else {
        if (Test-Path $repoDir) {
            # empty dir is fine for clone
            $left = @(Get-ChildItem -Force -LiteralPath $repoDir -ErrorAction SilentlyContinue)
            if ($left.Count -gt 0 -and -not (Test-Path (Join-Path $repoDir ".git"))) {
                Write-Bad "$repoDir is not empty and is not a git repo. Move it aside and re-run."
                exit 1
            }
        }
        Write-Step "Cloning $RepoUrl ($Branch) -> $repoDir"
        & git clone -b $Branch $RepoUrl $repoDir
        if ($LASTEXITCODE -ne 0) {
            Write-Bad "git clone failed (exit $LASTEXITCODE)"
            exit 1
        }
    }
}

Push-Location $repoDir
try {
    Write-Step "Fetching and checking out $Branch"
    & git fetch origin $Branch
    if ($LASTEXITCODE -ne 0) {
        Write-Bad "git fetch failed"
        exit 1
    }
    & git checkout -B $Branch "origin/$Branch"
    if ($LASTEXITCODE -ne 0) {
        Write-Bad "git checkout failed"
        exit 1
    }
    & git pull --ff-only origin $Branch 2>$null
    $head = (& git rev-parse --short HEAD | Out-String).Trim()
    Write-Ok "Checked out $Branch at $head in $repoDir"
}
finally {
    Pop-Location
}

$ps1 = Join-Path $repoDir "scripts\windows-sandbox-preflight.ps1"
$shWin = Join-Path $repoDir "scripts\windows-sandbox-preflight.sh"
$sandboxJson = Join-Path $repoDir ".cursor\sandbox.json"

foreach ($path in @($ps1, $shWin, $sandboxJson)) {
    if (-not (Test-Path -LiteralPath $path)) {
        Write-Bad "Missing required file: $path"
    }
    else {
        Write-Ok "Found $path"
    }
}
if ($fail -gt 0) {
    Write-Bad "Branch files are missing. Re-run after checking network/git access."
    exit 1
}

Write-Host ""
Write-Host "------------------------------------------------------------"
Write-Host " Running Windows preflight"
Write-Host "------------------------------------------------------------"
& powershell -NoProfile -ExecutionPolicy Bypass -File $ps1
$psExit = $LASTEXITCODE

$shWsl = $null
if ($shWin -match '^[A-Za-z]:') {
    $drive = $shWin.Substring(0, 1).ToLowerInvariant()
    $rest = $shWin.Substring(2) -replace '\\', '/'
    $shWsl = "/mnt/$drive$rest"
}
else {
    $shWsl = $shWin -replace '\\', '/'
}

Write-Host ""
Write-Host "------------------------------------------------------------"
Write-Host " Running WSL / Linux preflight"
Write-Host "------------------------------------------------------------"
Write-Step "WSL path: $shWsl"
$wslExit = 1
if (-not (Get-Command wsl -ErrorAction SilentlyContinue)) {
    Write-Host "[WARN] wsl.exe not found. Install with: wsl --install -d Ubuntu" -ForegroundColor Yellow
}
else {
    & wsl -e bash $shWsl
    $wslExit = $LASTEXITCODE
}

Write-Host ""
Write-Host "============================================================"
Write-Host " Open this folder in Cursor:"
Write-Host "   $repoDir"
Write-Host " Then set:"
Write-Host "   Approvals & Execution = Auto-review"
Write-Host "   Legacy Terminal Tool = Off"
Write-Host "============================================================"
Write-Host ""
Write-Host ("[RESULT] Windows preflight exit={0}  WSL preflight exit={1}" -f $psExit, $wslExit)

if (($psExit -ne 0) -or ($wslExit -ne 0)) {
    exit 1
}
Write-Ok "Preflight completed with exit code 0."
exit 0
