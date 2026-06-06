$ErrorActionPreference = 'Stop'

$runKey = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Run'
$name = 'ClaudeCodexGitAutosync'
$scriptPath = 'C:\Users\namma\.claude\git_autosync_launcher.ps1'
$psExe = 'C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe'
$value = "$psExe -NoProfile -WindowStyle Hidden -Command ""Get-Content '$scriptPath' -Raw | Invoke-Expression"""

New-ItemProperty -Path $runKey -Name $name -PropertyType String -Value $value -Force | Out-Null
Write-Output "Registered Run key: $name"
