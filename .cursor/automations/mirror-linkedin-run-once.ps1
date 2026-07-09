# One-shot mirror for a single LinkedIn run from GitHub (no -Topic param required on old scripts).
# Run in PowerShell 5.1 on Windows.

[CmdletBinding()]
param(
    [string]$Branch = "cursor/linkedin-figma-notion-claude-0681",
    [string]$Repo = "leejunha781/Codex",
    [string]$Date = "2026-07-09",
    [string]$Topic = "engineering-rca-evidence-structure",
    [string]$RepoRoot = "C:\Users\namma",
    [string]$MirrorTarget = "C:\Users\namma\Documents\Codex"
)

$ErrorActionPreference = "Stop"

$RunsPrefix = ".cursor/automations/daily-linkedin-marine-plm-post/runs"
$DestDir = Join-Path $MirrorTarget (Join-Path $Date $Topic)
$Files = @(
    "linkedin-post.md",
    "$Topic-infographic.png",
    "ready-for-posting.json"
)

New-Item -ItemType Directory -Force -Path $DestDir | Out-Null

if (Test-Path (Join-Path $RepoRoot ".git")) {
    Write-Host "Fetching origin/$Branch ..."
    git -C $RepoRoot fetch origin $Branch 2>&1 | Out-Null
}

$BaseRaw = "https://raw.githubusercontent.com/$Repo/$Branch"
$Copied = 0

foreach ($file in $Files) {
    $rel = "$RunsPrefix/$Date/$Topic/$file" -replace '\\', '/'
    $dest = Join-Path $DestDir $file
    $url = "$BaseRaw/$rel"

    Write-Host "GET $url"
    try {
        Invoke-WebRequest -Uri $url -UseBasicParsing -OutFile $dest
        Write-Host "OK   $dest"
        $Copied++
    } catch {
        # Fallback: git show if raw URL fails (e.g. PNG LFS)
        if (Test-Path (Join-Path $RepoRoot ".git")) {
            $gitPath = $rel
            $gitRef = "origin/${Branch}:$gitPath"
            Write-Host "TRY  git show $gitRef"
            cmd.exe /c "git -C `"$RepoRoot`" show `"$gitRef`" > `"$dest`"" | Out-Null
            if (Test-Path $dest) {
                Write-Host "OK   $dest (git show)"
                $Copied++
            } else {
                Write-Warning "FAIL $file"
            }
        } else {
            Write-Warning "FAIL $file — $($_.Exception.Message)"
        }
    }
}

$manifest = [ordered]@{
    topic             = $Topic
    date              = $Date
    status            = "mirrored_ready_to_post"
    windowsMirrorPath = $DestDir
    postPath          = Join-Path $DestDir "linkedin-post.md"
    imagePath         = Join-Path $DestDir "$Topic-infographic.png"
    linkedInApp       = "Windows LinkedIn app only"
    autoPost          = $true
}
$manifest | ConvertTo-Json -Depth 4 | Set-Content -Path (Join-Path $DestDir "ready-for-posting.json") -Encoding UTF8

Write-Host ""
Write-Host "Mirror complete: copied=$Copied dest=$DestDir"
if ($Copied -ge 2) {
    Write-Host "READY — open LinkedIn Windows app and post from:"
    Write-Host "  Post:  $(Join-Path $DestDir 'linkedin-post.md')"
    Write-Host "  Image: $(Join-Path $DestDir "$Topic-infographic.png")"
} else {
    Write-Host "WARN — some files missing. Check branch $Branch on GitHub."
    exit 1
}
