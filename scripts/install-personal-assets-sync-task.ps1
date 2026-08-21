param(
    [string]$RepositoryRoot = (Split-Path -Parent $PSScriptRoot),
    [int]$IntervalMinutes = 15
)

$ErrorActionPreference = 'Stop'
$repoRoot = [IO.Path]::GetFullPath($RepositoryRoot)
$syncScript = Join-Path $repoRoot 'scripts\sync-personal-assets.ps1'
if (-not (Test-Path -LiteralPath $syncScript)) { throw "Missing sync script: $syncScript" }

$action = New-ScheduledTaskAction -Execute 'powershell.exe' -Argument (
    "-NoProfile -ExecutionPolicy Bypass -File `"$syncScript`" -RepositoryRoot `"$repoRoot`" -Push"
)
$trigger = New-ScheduledTaskTrigger -Once -At (Get-Date).AddMinutes(1) `
    -RepetitionInterval (New-TimeSpan -Minutes $IntervalMinutes)
$settings = New-ScheduledTaskSettingsSet -MultipleInstances IgnoreNew -StartWhenAvailable
Register-ScheduledTask -TaskName 'Codex-Personal-Assets-GitHub-Sync' -Action $action `
    -Trigger $trigger -Settings $settings -Description 'Sync personal Codex skills and Sites sources to GitHub.' -Force
