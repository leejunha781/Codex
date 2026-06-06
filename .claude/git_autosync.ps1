# git_autosync.ps1
# Claude + Codex 작업 폴더 변경 감지 -> 자동 commit + push to GitHub
# 감시 대상: C:\Users\namma\.claude\  +  C:\Users\namma\.codex\skills\  +  C:\Users\namma\.codex\automations\

$repoRoot  = "C:\Users\namma"
$logFile   = "C:\Users\namma\.claude\cache\git-autosync\autosync.log"
$debounce  = 45   # 마지막 변경 후 이 초 동안 새 변경 없으면 commit
$watchPaths = @(
    "C:\Users\namma\.claude",
    "C:\Users\namma\.codex\skills",
    "C:\Users\namma\.codex\automations"
)

# 로그 폴더 준비
New-Item -ItemType Directory -Force -Path (Split-Path $logFile) | Out-Null

function Write-Log($msg) {
    $line = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')  $msg"
    Add-Content $logFile $line
}

Write-Log "=== autosync started. watching: $($watchPaths -join ', ') ==="

# 변경 타이머 공유 변수
$script:lastChange = $null
$script:pending    = $false

# FileSystemWatcher 생성
$watchers = @()
foreach ($p in $watchPaths) {
    if (-not (Test-Path $p)) { continue }
    $w = New-Object System.IO.FileSystemWatcher
    $w.Path                  = $p
    $w.IncludeSubdirectories = $true
    $w.NotifyFilter          = [System.IO.NotifyFilters]::LastWrite -bor
                               [System.IO.NotifyFilters]::FileName  -bor
                               [System.IO.NotifyFilters]::DirectoryName
    # 제외 패턴 (git 내부, 렌더 이미지 등)
    $w.Filter = "*"
    $w.EnableRaisingEvents = $true

    $action = {
        $path = $Event.SourceEventArgs.FullPath
        # git 내부 변경이나 이미지/바이너리는 무시
        if ($path -match '\\\.git\\' -or
            $path -match '\.(png|jpg|jpeg|pdf|docx|pptx|zip|sqlite|wal|shm|tmp)$' -or
            $path -match '\\cache\\' -or
            $path -match 'autosync\.log') { return }
        $script:lastChange = Get-Date
        $script:pending    = $true
    }

    Register-ObjectEvent $w "Changed" -Action $action | Out-Null
    Register-ObjectEvent $w "Created" -Action $action | Out-Null
    Register-ObjectEvent $w "Deleted" -Action $action | Out-Null
    Register-ObjectEvent $w "Renamed" -Action $action | Out-Null
    $watchers += $w
    Write-Log "watching: $p"
}

Write-Log "debounce = ${debounce}s  |  loop started"

# 메인 루프
while ($true) {
    Start-Sleep -Seconds 5

    if ($script:pending -and $script:lastChange) {
        $elapsed = ((Get-Date) - $script:lastChange).TotalSeconds
        if ($elapsed -ge $debounce) {
            $script:pending    = $false
            $script:lastChange = $null

            Push-Location $repoRoot
            try {
                # 변경 파일 스테이징
                git add ".claude/" ".codex/skills/" ".codex/automations/" 2>$null

                # 변경 있는지 확인
                $diff = git diff --cached --name-only 2>$null
                if ($diff) {
                    $ts  = Get-Date -Format "yyyy-MM-dd HH:mm"
                    $msg = "auto-sync: $ts ($($diff.Count) files)"
                    git commit -m $msg 2>$null
                    $push = git push 2>&1
                    Write-Log "PUSHED  $msg  | $push"
                }
            } catch {
                Write-Log "ERROR: $_"
            } finally {
                Pop-Location
            }
        }
    }
}
