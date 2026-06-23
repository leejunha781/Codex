$ErrorActionPreference = 'Stop'

$syncScript = 'C:\Users\namma\.claude\ensure_git_autosync_connection.ps1'

if (-not (Test-Path -LiteralPath $syncScript -PathType Leaf)) {
    throw "Sync script not found: $syncScript"
}

Get-Content $syncScript -Raw | Invoke-Expression
