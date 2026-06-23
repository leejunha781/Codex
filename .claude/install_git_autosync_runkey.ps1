$ErrorActionPreference = 'Stop'

$runKey = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Run'
$name = 'ClaudeCodexGitAutosync'

if (Get-ItemProperty -Path $runKey -Name $name -ErrorAction SilentlyContinue) {
    Remove-ItemProperty -Path $runKey -Name $name
    Write-Output "Removed Run key: $name"
} else {
    Write-Output "Run key already absent: $name"
}

Write-Output "Use Codex Automations entry 'Manual Claude Codex Sync' or run C:\Users\namma\.claude\start_claude_codex_sync.ps1 manually."
