$ErrorActionPreference = 'Stop'

$repoRoot = 'C:\Users\namma'
$originUrl = 'https://github.com/leejunha781/Codex.git'
$launcher = 'C:\Users\namma\.claude\git_autosync_launcher.ps1'
$logFile = 'C:\Users\namma\.claude\cache\git-autosync\autosync.log'

New-Item -ItemType Directory -Force -Path (Split-Path $logFile) | Out-Null

function Write-Log {
    param([string]$Message)
    Add-Content $logFile "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')  connection: $Message"
}

if (-not (Test-Path -LiteralPath $repoRoot -PathType Container)) {
    throw "Repo root not found: $repoRoot"
}

if (-not (Test-Path -LiteralPath (Join-Path $repoRoot '.git') -PathType Container)) {
    git -C $repoRoot init | Out-Null
    Write-Log "initialized root repo at $repoRoot"
}

$currentOrigin = git -C $repoRoot remote get-url origin 2>$null
if ($LASTEXITCODE -ne 0 -or [string]::IsNullOrWhiteSpace($currentOrigin)) {
    git -C $repoRoot remote add origin $originUrl
    Write-Log "added origin $originUrl"
} elseif ($currentOrigin.Trim() -ne $originUrl) {
    git -C $repoRoot remote set-url origin $originUrl
    Write-Log "updated origin to $originUrl"
} else {
    Write-Log "origin already set to $originUrl"
}

Get-Content $launcher -Raw | Invoke-Expression
