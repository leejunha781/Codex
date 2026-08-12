# One-shot mirror for a single LinkedIn run from GitHub (PowerShell 5.1 safe).
# Does NOT require git fetch — downloads directly from raw.githubusercontent.com.

[CmdletBinding()]
param(
    [string]$Branch = "memory",
    [string]$Repo = "leejunha781/Codex",
    [string]$Date = "2026-07-09",
    [string]$Topic = "engineering-rca-evidence-structure",
    [string]$RepoRoot = "C:\Users\namma",
    [string]$MirrorTarget = "C:\Users\namma\Documents\Codex"
)

$ErrorActionPreference = "Stop"

$RunsPrefix = ".cursor/automations/daily-linkedin-marine-plm-post/runs"
$DestDir = Join-Path $MirrorTarget (Join-Path $Date $Topic)
$BaseRaw = "https://raw.githubusercontent.com/$Repo/$Branch"

New-Item -ItemType Directory -Force -Path $DestDir | Out-Null

function Get-RemoteFile {
    param(
        [string]$RelativePath,
        [string]$DestPath
    )

    $url = "$BaseRaw/$($RelativePath -replace '\\', '/')"
    Write-Host "GET $url"

    $prev = $ErrorActionPreference
    $ErrorActionPreference = "Continue"
    try {
        Invoke-WebRequest -Uri $url -UseBasicParsing -OutFile $DestPath -ErrorAction Stop
        if ((Test-Path $DestPath) -and ((Get-Item $DestPath).Length -gt 0)) {
            Write-Host "OK   $DestPath"
            return $true
        }
    } catch {
        Write-Host "WARN raw download failed: $($_.Exception.Message)"
    } finally {
        $ErrorActionPreference = $prev
    }

    if (-not (Test-Path (Join-Path $RepoRoot ".git"))) {
        return $false
    }

    $gitRef = "origin/${Branch}:$($RelativePath -replace '\\', '/')"
    Write-Host "TRY  git show $gitRef"
    $quotedRoot = '"' + $RepoRoot + '"'
    cmd.exe /c "git -C $quotedRoot fetch origin $Branch 2>nul" | Out-Null

    # Binary-safe export: cmd `>` corrupts PNGs via text-mode CRLF rewriting.
    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = "git"
    $psi.Arguments = "-C `"$RepoRoot`" show `"$gitRef`""
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError = $true
    $psi.UseShellExecute = $false
    $psi.CreateNoWindow = $true
    $proc = [System.Diagnostics.Process]::Start($psi)
    $fs = [System.IO.File]::Create($DestPath)
    try {
        $proc.StandardOutput.BaseStream.CopyTo($fs)
    } finally {
        $fs.Dispose()
    }
    $proc.WaitForExit()

    if ($proc.ExitCode -eq 0 -and (Test-Path $DestPath) -and ((Get-Item $DestPath).Length -gt 0)) {
        Write-Host "OK   $DestPath (git show)"
        return $true
    }

    Write-Warning "FAIL $RelativePath"
    return $false
}

$Copied = 0
$postRel = "$RunsPrefix/$Date/$Topic/linkedin-post.md"
$imageRel = "$RunsPrefix/$Date/$Topic/$Topic-infographic.png"
$postDest = Join-Path $DestDir "linkedin-post.md"
$imageDest = Join-Path $DestDir "$Topic-infographic.png"

if (Get-RemoteFile -RelativePath $postRel -DestPath $postDest) { $Copied++ }
if (Get-RemoteFile -RelativePath $imageRel -DestPath $imageDest) { $Copied++ }

$manifest = [ordered]@{
    topic             = $Topic
    date              = $Date
    status            = "mirrored_ready_to_post"
    windowsMirrorPath = $DestDir
    postPath          = $postDest
    imagePath         = $imageDest
    linkedInApp       = "Windows LinkedIn app only"
    autoPost          = $true
    branch            = $Branch
}
$manifest | ConvertTo-Json -Depth 4 | Set-Content -Path (Join-Path $DestDir "ready-for-posting.json") -Encoding UTF8

Write-Host ""
Write-Host "Mirror complete: copied=$Copied dest=$DestDir"
if ($Copied -ge 2) {
    Write-Host "READY — open LinkedIn Windows app and post from:"
    Write-Host "  Post:  $postDest"
    Write-Host "  Image: $imageDest"
    exit 0
}

Write-Host "WARN — some files missing. Check branch $Branch on GitHub."
exit 1
