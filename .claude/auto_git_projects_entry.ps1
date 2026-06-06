param(
    [switch]$RunOnce,
    [string]$TargetPath
)

$ErrorActionPreference = 'Stop'

$root = 'C:\Users\namma\.claude\projects\C--Users-namma--claude'
$stateDir = 'C:\Users\namma\.claude\cache\git-auto-init'
$logPath = Join-Path $stateDir 'watcher.log'
$pidPath = Join-Path $stateDir 'watcher.pid'
$excludeNames = @('memory')

New-Item -ItemType Directory -Force -Path $stateDir | Out-Null

function Write-Log {
    param([string]$Message)
    $stamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    Add-Content -Path $logPath -Value "$stamp $Message"
}

function Is-TrackedProjectDir {
    param([string]$Path)
    if (-not (Test-Path -LiteralPath $Path -PathType Container)) { return $false }
    $name = Split-Path -Leaf $Path
    if ($excludeNames -contains $name) { return $false }
    return $true
}

function Ensure-GitRepo {
    param([string]$ProjectPath)

    if (-not (Is-TrackedProjectDir -Path $ProjectPath)) { return }
    $gitDir = Join-Path $ProjectPath '.git'
    if (Test-Path -LiteralPath $gitDir) {
        Write-Log "skip existing repo: $ProjectPath"
        return
    }

    try {
        & git -C $ProjectPath init -b main | Out-Null
        if ($LASTEXITCODE -eq 0) {
            Write-Log "initialized repo: $ProjectPath"
        } else {
            Write-Log "git init failed ($LASTEXITCODE): $ProjectPath"
        }
    } catch {
        Write-Log "git init exception: $ProjectPath :: $($_.Exception.Message)"
    }
}

function Ensure-ExistingProjects {
    Get-ChildItem -LiteralPath $root -Directory | ForEach-Object {
        Ensure-GitRepo -ProjectPath $_.FullName
    }
}

if ($RunOnce) {
    Ensure-GitRepo -ProjectPath $TargetPath
    return
}

if (-not (Test-Path -LiteralPath $root -PathType Container)) {
    throw "Root not found: $root"
}

if (Test-Path -LiteralPath $pidPath) {
    try {
        $existingPid = [int](Get-Content -LiteralPath $pidPath -ErrorAction Stop)
        $existing = Get-Process -Id $existingPid -ErrorAction SilentlyContinue
        if ($existing) {
            Write-Log "watcher already running pid=$existingPid"
            return
        }
    } catch {
        Write-Log "stale pid file ignored"
    }
}

Set-Content -Path $pidPath -Value $PID
Write-Log "watcher start pid=$PID"

Ensure-ExistingProjects

$watcher = New-Object System.IO.FileSystemWatcher
$watcher.Path = $root
$watcher.Filter = '*'
$watcher.IncludeSubdirectories = $false
$watcher.NotifyFilter = [System.IO.NotifyFilters]'DirectoryName, CreationTime'
$watcher.EnableRaisingEvents = $true

$action = {
    $path = $Event.SourceEventArgs.FullPath
    if (Test-Path -LiteralPath $path -PathType Container) {
        Start-Sleep -Milliseconds 500
        & 'C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe' -NoProfile -WindowStyle Hidden -File 'C:\Users\namma\.claude\auto_git_projects_entry.ps1' -RunOnce -TargetPath $path | Out-Null
    }
}

Register-ObjectEvent -InputObject $watcher -EventName Created -SourceIdentifier 'ClaudeProjectGitInit.Created' -Action $action | Out-Null

try {
    while ($true) {
        $evt = Wait-Event -SourceIdentifier 'ClaudeProjectGitInit.Created' -Timeout 60
        if ($evt) {
            Remove-Event -EventIdentifier $evt.EventIdentifier -ErrorAction SilentlyContinue
        }
    }
} finally {
    Unregister-Event -SourceIdentifier 'ClaudeProjectGitInit.Created' -ErrorAction SilentlyContinue
    $watcher.EnableRaisingEvents = $false
    $watcher.Dispose()
    Write-Log "watcher stop pid=$PID"
}
