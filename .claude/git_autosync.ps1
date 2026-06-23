# git_autosync.ps1
# Claude + Codex 작업 폴더 변경 감지 -> 자동 commit + push to GitHub
# 감시 대상: Claude + Codex 설정, 프로젝트, chat/code/cowork 세션 텍스트

$repoRoot  = "C:\Users\namma"
$logFile   = "C:\Users\namma\.claude\cache\git-autosync\autosync.log"
$debounce  = 45   # 마지막 변경 후 이 초 동안 새 변경 없으면 commit
$watchPaths = @(
    "C:\Users\namma\.claude",
    "C:\Users\namma\.codex\skills",
    "C:\Users\namma\.codex\automations",
    "C:\Users\namma\.codex\sessions",
    "C:\Users\namma\.codex\archived_sessions",
    "C:\Users\namma\.codex\attachments",
    "C:\Users\namma\.codex\session_index.jsonl",
    "C:\Users\namma\.codex\external_agent_session_imports.json",
    "C:\Users\namma\.codex\process_manager\chat_processes.json",
    "C:\Users\namma\.gitignore",
    "C:\Users\namma\.gitattributes"
)
$watchSpecs = @(
    ".claude/",
    ".codex/skills/",
    ".codex/automations/",
    ".codex/sessions/",
    ".codex/archived_sessions/",
    ".codex/attachments/",
    ".codex/session_index.jsonl",
    ".codex/external_agent_session_imports.json",
    ".codex/process_manager/chat_processes.json",
    ".gitignore",
    ".gitattributes"
)
$repoRootFull = (Resolve-Path -LiteralPath $repoRoot).ProviderPath.TrimEnd('\')

# 로그 폴더 준비
New-Item -ItemType Directory -Force -Path (Split-Path $logFile) | Out-Null

function Write-Log($msg) {
    $line = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')  $msg"
    Add-Content $logFile $line
}

Write-Log "=== autosync started. watching: $($watchPaths -join ', ') ==="

foreach ($p in $watchPaths) {
    if (Test-Path $p) {
        Write-Log "watching: $p"
    } else {
        Write-Log "skip missing path: $p"
    }
}

function Get-ActiveWatchSpecs {
    $active = @()
    foreach ($spec in $watchSpecs) {
        $relative = $spec.TrimEnd("/")
        $fullPath = Join-Path $repoRoot $relative
        if (Test-Path $fullPath) {
            $active += $spec
        }
    }
    return $active
}

function Test-IsNestedGitRepoPath {
    param([string]$Path)

    $normalized = ($Path -replace '\\','/').Trim('/')
    if ([string]::IsNullOrWhiteSpace($normalized)) { return $false }

    $parts = @($normalized -split '/')
    for ($i = $parts.Count; $i -ge 1; $i--) {
        $candidateRel = ($parts[0..($i - 1)] -join '/')
        $candidateFull = Join-Path $repoRoot ($candidateRel -replace '/', [System.IO.Path]::DirectorySeparatorChar)
        $resolved = Resolve-Path -LiteralPath $candidateFull -ErrorAction SilentlyContinue
        if (-not $resolved) { continue }

        $candidateRoot = $resolved.ProviderPath.TrimEnd('\')
        if ($candidateRoot -ieq $repoRootFull) { continue }

        if (Test-Path -LiteralPath (Join-Path $candidateRoot '.git')) {
            return $true
        }
    }

    return $false
}

function Should-IgnoreGitPath {
    param([string]$Path)

    $isClaudeChatArtifact = $Path -match '^\.claude/projects/C--Users-namma--claude/[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}/(subagents|tool-results)/'

    return (
        $Path -match '(^|/)\.git/' -or
        ((Test-IsNestedGitRepoPath -Path $Path) -and -not $isClaudeChatArtifact) -or
        (($Path -match '^\.claude/projects/C--Users-namma--claude/[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}(/|$)') -and -not $isClaudeChatArtifact) -or
        $Path -match '(^|/)cache/' -or
        $Path -match 'autosync\.log$' -or
        $Path -match '\.(png|jpg|jpeg|pdf|docx|pptx|zip|sqlite|wal|shm|tmp)$'
    )
}

function Get-ChangePaths {
    $paths = @()
    $activeSpecs = @(Get-ActiveWatchSpecs)
    if ($activeSpecs.Count -eq 0) { return $paths }

    Push-Location $repoRoot
    try {
        $statusLines = @(git status --porcelain -- @activeSpecs 2>$null)
        foreach ($line in $statusLines) {
            if (-not $line -or $line.Length -lt 4) { continue }
            $pathPart = $line.Substring(3)

            if ($pathPart -match ' -> ') {
                foreach ($renamedPath in ($pathPart -split ' -> ')) {
                    if ($renamedPath -and -not (Should-IgnoreGitPath -Path $renamedPath)) {
                        $paths += $renamedPath
                    }
                }
            } elseif (-not (Should-IgnoreGitPath -Path $pathPart)) {
                $paths += $pathPart
            }
        }
    } finally {
        Pop-Location
    }

    return @($paths | Sort-Object -Unique)
}

function Invoke-GitSync {
    param([string[]]$Paths)

    if (-not $Paths -or $Paths.Count -eq 0) { return }

    Push-Location $repoRoot
    try {
        git add --all -- @Paths 2>$null
        if ($LASTEXITCODE -ne 0) {
            Write-Log "ERROR: git add failed for $($Paths.Count) paths"
            return
        }

        $staged = @(git diff --cached --name-only 2>$null)
        if ($staged.Count -eq 0) { return }

        $ts  = Get-Date -Format "yyyy-MM-dd HH:mm"
        $msg = "auto-sync: $ts ($($staged.Count) files)"
        $commit = git commit -m $msg 2>&1
        if ($LASTEXITCODE -ne 0) {
            Write-Log "ERROR: git commit failed | $commit"
            return
        }

        $push = git push 2>&1
        if ($LASTEXITCODE -eq 0) {
            Write-Log "PUSHED  $msg  | $push"
        } else {
            Write-Log "ERROR: git push failed for $msg | $push"
        }
    } catch {
        Write-Log "ERROR: $_"
    } finally {
        Pop-Location
    }
}

Write-Log "debounce = ${debounce}s  |  polling loop started"

$lastSignature = $null
$lastChange = $null

# 메인 루프
while ($true) {
    Start-Sleep -Seconds 5

    $changePaths = @(Get-ChangePaths)
    if ($changePaths.Count -eq 0) {
        $lastSignature = $null
        $lastChange = $null
        continue
    }

    $signature = ($changePaths | Sort-Object) -join "`n"
    if ($signature -ne $lastSignature) {
        $lastSignature = $signature
        $lastChange = Get-Date
        Write-Log "pending changes detected: $($changePaths.Count) paths"
        continue
    }

    $elapsed = ((Get-Date) - $lastChange).TotalSeconds
    if ($elapsed -ge $debounce) {
        Invoke-GitSync -Paths $changePaths
        $lastSignature = $null
        $lastChange = $null
    }
}
