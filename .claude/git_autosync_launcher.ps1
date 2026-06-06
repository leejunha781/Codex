# git_autosync_launcher.ps1
# 로그인 시 autosync를 숨김 창으로 실행하는 런처
# 중복 실행 방지: 이미 autosync가 돌고 있으면 새로 띄우지 않음

$lockFile = "C:\Users\namma\.claude\cache\git-autosync\autosync.pid"
$script   = "C:\Users\namma\.claude\git_autosync.ps1"
$logFile  = "C:\Users\namma\.claude\cache\git-autosync\autosync.log"

New-Item -ItemType Directory -Force -Path (Split-Path $lockFile) | Out-Null

# 기존 프로세스 살아있는지 확인
if (Test-Path $lockFile) {
    $existingPid = Get-Content $lockFile -ErrorAction SilentlyContinue
    if ($existingPid -and (Get-Process -Id $existingPid -ErrorAction SilentlyContinue)) {
        Add-Content $logFile "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')  launcher: already running (pid=$existingPid), skip"
        exit 0
    }
}

# 새 프로세스로 실행
$proc = Start-Process powershell.exe `
    -ArgumentList "-NonInteractive -WindowStyle Hidden -Command `"Get-Content '$script' -Raw | Invoke-Expression`"" `
    -PassThru -WindowStyle Hidden

$proc.Id | Out-File $lockFile -Encoding ascii
Add-Content $logFile "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')  launcher: started pid=$($proc.Id)"
