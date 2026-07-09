# Mirror Daily LinkedIn run artifacts from repo to local Windows output folder.
# Cloud Cursor runs write to repo runs/ only; this script copies to Documents\Codex.
# Windows mirror path: C:\Users\namma\Documents\Codex\YYYY-MM-DD\<topic-slug>\
# Run from PowerShell 5.1 on Windows.

[CmdletBinding()]
param(
    [string]$RepoRoot = "C:\Users\namma",
    [string]$MirrorTarget = "C:\Users\namma\Documents\Codex",
    [string]$RunsRelative = ".cursor/automations/daily-linkedin-marine-plm-post/runs",
    [string]$Date,
    [string]$Topic,
    [switch]$AllDates,
    [switch]$Pull,
    [switch]$IncludeRemoteBranches,
    [switch]$DryRun
)

$ErrorActionPreference = "Stop"

$WindowsMirrorPathTemplate = "C:\Users\namma\Documents\Codex\YYYY-MM-DD\<topic-slug>\"

$RunsRelativeNormalized = ($RunsRelative -replace '\\', '/').Trim('/')
$RunsSource = Join-Path $RepoRoot ($RunsRelativeNormalized -replace '/', [System.IO.Path]::DirectorySeparatorChar)
$LogDir = Join-Path $RepoRoot ".cursor\automations\cache\linkedin-mirror"
$LogFile = Join-Path $LogDir "mirror.log"

New-Item -ItemType Directory -Force -Path $LogDir | Out-Null

function Write-MirrorLog {
    param([string]$Message)
    $line = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')  $Message"
    Add-Content -Path $LogFile -Value $line
    Write-Host $Message
}

function Test-IsRunArtifactName {
    param([string]$FileName)
    return (
        $FileName -eq "linkedin-post.md" -or
        $FileName -eq "ready-for-posting.json" -or
        $FileName -like "*-infographic.png"
    )
}

function Invoke-GitCommand {
    param(
        [string]$Root,
        [string[]]$GitArgs
    )

    $output = & git -C $Root @GitArgs 2>&1
    $exitCode = $LASTEXITCODE
    return @{
        Output   = $output
        ExitCode = $exitCode
    }
}

function Invoke-GitPull {
    param([string]$Root)

    if (-not (Test-Path (Join-Path $Root ".git"))) {
        Write-MirrorLog "WARN: git repo not found at $Root; skip pull"
        return
    }

    $branch = (Invoke-GitCommand -Root $Root -GitArgs @("rev-parse", "--abbrev-ref", "HEAD")).Output
    $branch = ($branch | Select-Object -First 1).ToString().Trim()
    if ([string]::IsNullOrWhiteSpace($branch) -or $branch -eq "HEAD") {
        Write-MirrorLog "WARN: detached HEAD at $Root; fetch only"
        $fetch = Invoke-GitCommand -Root $Root -GitArgs @("fetch", "origin", "--prune")
        if ($fetch.ExitCode -ne 0) {
            Write-MirrorLog "WARN: git fetch failed | $($fetch.Output -join ' ')"
        }
        return
    }

    $fetch = Invoke-GitCommand -Root $Root -GitArgs @("fetch", "origin", "--prune")
    if ($fetch.ExitCode -ne 0) {
        Write-MirrorLog "WARN: git fetch failed | $($fetch.Output -join ' ')"
    }

    $remoteRef = Invoke-GitCommand -Root $Root -GitArgs @("rev-parse", "--verify", "origin/$branch")
    if ($remoteRef.ExitCode -eq 0) {
        $pull = Invoke-GitCommand -Root $Root -GitArgs @("pull", "--rebase", "--autostash", "origin", $branch)
        if ($pull.ExitCode -eq 0) {
            Write-MirrorLog "Pulled origin/$branch"
        } else {
            Write-MirrorLog "WARN: pull failed | $($pull.Output -join ' ')"
        }
    }

    $lfs = Invoke-GitCommand -Root $Root -GitArgs @("lfs", "pull")
    if ($lfs.ExitCode -ne 0) {
        Write-MirrorLog "WARN: git lfs pull failed | $($lfs.Output -join ' ')"
    }
}

function Get-LocalRunArtifacts {
    param(
        [string]$SourceRoot,
        [string]$FilterDate,
        [string]$FilterTopic,
        [bool]$IncludeAllDates
    )

    $artifacts = @{}
    if (-not (Test-Path $SourceRoot)) {
        return $artifacts
    }

    $dateDirs = Get-ChildItem -Path $SourceRoot -Directory -ErrorAction SilentlyContinue
    foreach ($dateDir in $dateDirs) {
        if ($dateDir.Name -notmatch '^\d{4}-\d{2}-\d{2}$') { continue }
        if (-not $IncludeAllDates -and $FilterDate -and $dateDir.Name -ne $FilterDate) { continue }

        $topicDirs = Get-ChildItem -Path $dateDir.FullName -Directory -ErrorAction SilentlyContinue
        foreach ($topicDir in $topicDirs) {
            if ($FilterTopic -and $topicDir.Name -ne $FilterTopic) { continue }

            $files = Get-ChildItem -Path $topicDir.FullName -File -ErrorAction SilentlyContinue
            foreach ($file in $files) {
                if (-not (Test-IsRunArtifactName -FileName $file.Name)) { continue }
                $key = "$($dateDir.Name)/$($topicDir.Name)/$($file.Name)"
                $artifacts[$key] = @{
                    SourceKind = "local"
                    SourcePath = $file.FullName
                    UpdatedAt  = $file.LastWriteTimeUtc
                }
            }
        }
    }

    return $artifacts
}

function Get-RemoteRunArtifacts {
    param(
        [string]$Root,
        [string]$RunsPrefix,
        [string]$FilterDate,
        [string]$FilterTopic,
        [bool]$IncludeAllDates
    )

    $artifacts = @{}
    if (-not (Test-Path (Join-Path $Root ".git"))) {
        return $artifacts
    }

    git -C $Root fetch origin --prune 2>&1 | Out-Null

    $branchPatterns = @(
        "origin/cursor/linkedin-figma-notion-claude-0681",
        "origin/cursor/fix-linkedin-local-mirror-0681",
        "origin/cursor/daily-linkedin-test-non-plm-0681",
        "origin/cursor/daily-linkedin*",
        "origin/cursor/linkedin-daily*",
        "origin/cursor/linkedin-daily-automation-42c5",
        "origin/main",
        "origin/master",
        "origin/memory"
    )

    $branches = @()
    foreach ($pattern in $branchPatterns) {
        $branches += @(git -C $Root branch -r --list $pattern 2>$null)
    }
    $branches = @($branches | Where-Object { $_ } | Sort-Object -Unique)

    foreach ($branch in $branches) {
        $branch = $branch.Trim()
        $commitIso = (git -C $Root log -1 --format=%cI $branch 2>$null).Trim()
        if ([string]::IsNullOrWhiteSpace($commitIso)) { continue }

        $commitTime = [datetime]::Parse($commitIso).ToUniversalTime()
        $files = @(git -C $Root ls-tree -r --name-only $branch -- "$RunsPrefix/" 2>$null)
        foreach ($file in $files) {
            $normalized = ($file -replace '\\', '/')
            if ($normalized -notmatch 'runs/(\d{4}-\d{2}-\d{2})/([^/]+)/([^/]+)$') { continue }

            $runDate = $matches[1]
            $topicSlug = $matches[2]
            $fileName = $matches[3]

            if (-not (Test-IsRunArtifactName -FileName $fileName)) { continue }
            if (-not $IncludeAllDates -and $FilterDate -and $runDate -ne $FilterDate) { continue }
            if ($FilterTopic -and $topicSlug -ne $FilterTopic) { continue }

            $key = "$runDate/$topicSlug/$fileName"
            if (-not $artifacts.ContainsKey($key) -or $commitTime -gt $artifacts[$key].UpdatedAt) {
                $artifacts[$key] = @{
                    SourceKind = "remote"
                    Branch     = $branch
                    GitPath    = $normalized
                    UpdatedAt  = $commitTime
                }
            }
        }
    }

    return $artifacts
}

function Export-RemoteGitFile {
    param(
        [string]$Root,
        [string]$Branch,
        [string]$GitPath,
        [string]$DestPath
    )

    New-Item -ItemType Directory -Force -Path (Split-Path $DestPath -Parent) | Out-Null
    $gitRef = "$Branch`:$GitPath"
    $quotedRoot = '"' + $Root + '"'
    $quotedRef = '"' + $gitRef + '"'
    $quotedDest = '"' + $DestPath + '"'
    cmd.exe /c "git -C $quotedRoot show $quotedRef > $quotedDest" | Out-Null
    if ($LASTEXITCODE -ne 0 -or -not (Test-Path $DestPath)) {
        throw "Failed to export $gitRef to $DestPath"
    }
}

function Merge-ArtifactMaps {
    param(
        [hashtable]$Primary,
        [hashtable]$Secondary
    )

    $merged = @{}
    foreach ($entry in $Primary.GetEnumerator()) {
        $merged[$entry.Key] = $entry.Value
    }
    foreach ($entry in $Secondary.GetEnumerator()) {
        $key = $entry.Key
        if (-not $merged.ContainsKey($key) -or $entry.Value.UpdatedAt -gt $merged[$key].UpdatedAt) {
            $merged[$key] = $entry.Value
        }
    }
    return $merged
}

function Get-WindowsMirrorFolder {
    param(
        [string]$TargetRoot,
        [string]$RunDate,
        [string]$TopicSlug
    )

    return Join-Path $TargetRoot (Join-Path $RunDate $TopicSlug)
}

function Test-PngHasC2paChunk {
    param([string]$Path)

    $bytes = [System.IO.File]::ReadAllBytes($Path)
    $text = [System.Text.Encoding]::ASCII.GetString($bytes)
    return ($text -match 'caBX' -or $text -match 'c2pa' -or $text -match 'jumb')
}

function Remove-C2paFromLinkedInPng {
    param([string]$Path)

    if (-not (Test-Path -LiteralPath $Path)) { return }
    if ($Path -notmatch '\.png$') { return }
    if (-not (Test-PngHasC2paChunk -Path $Path)) { return }

    Add-Type -AssemblyName System.Drawing
    $image = $null
    try {
        $image = [System.Drawing.Image]::FromFile($Path)
        $tempPath = "$Path.tmp.png"
        $image.Save($tempPath, [System.Drawing.Imaging.ImageFormat]::Png)
        $image.Dispose()
        $image = $null
        Move-Item -LiteralPath $tempPath -Destination $Path -Force
        Write-MirrorLog "Stripped C2PA metadata from $Path"
    }
    finally {
        if ($null -ne $image) { $image.Dispose() }
    }
}

function Write-ReadyForPostingManifest {
    param(
        [string]$TargetRoot,
        [string]$RunDate,
        [string]$TopicSlug,
        [bool]$IsDryRun
    )

    $destDir = Get-WindowsMirrorFolder -TargetRoot $TargetRoot -RunDate $RunDate -TopicSlug $TopicSlug
    $postPath = Join-Path $destDir "linkedin-post.md"
    $imageCandidates = @(Get-ChildItem -Path $destDir -Filter "*-infographic.png" -File -ErrorAction SilentlyContinue)
    $imagePath = if ($imageCandidates.Count -gt 0) { $imageCandidates[0].FullName } else { Join-Path $destDir "$TopicSlug-infographic.png" }

    $manifest = [ordered]@{
        topic                 = $TopicSlug
        date                  = $RunDate
        status                = "mirrored_ready_to_post"
        windowsMirrorPath     = $destDir
        windowsMirrorTemplate = $WindowsMirrorPathTemplate
        postPath              = $postPath
        imagePath             = $imagePath
        linkedInApp           = "Windows LinkedIn app only"
        autoPost              = $true
        mirrorScript          = "C:\Users\namma\.cursor\automations\mirror-linkedin-runs.ps1"
    }

    $manifestPath = Join-Path $destDir "ready-for-posting.json"
    if ($IsDryRun) {
        Write-MirrorLog "[dry-run] would write manifest $manifestPath"
        return $manifest
    }

    if (-not (Test-Path $postPath)) {
        Write-MirrorLog "WARN: skip manifest for $RunDate/$TopicSlug (missing linkedin-post.md)"
        return $null
    }

    $manifest | ConvertTo-Json -Depth 4 | Set-Content -Path $manifestPath -Encoding UTF8
    Write-MirrorLog "Wrote manifest $manifestPath"
    return $manifest
}

if ($Pull) {
    Write-MirrorLog "Pull requested for $RepoRoot"
    Invoke-GitPull -Root $RepoRoot
}

$filterDate = $Date
if (-not $AllDates -and [string]::IsNullOrWhiteSpace($filterDate)) {
    $filterDate = (Get-Date).ToString("yyyy-MM-dd")
}

$localArtifacts = Get-LocalRunArtifacts -SourceRoot $RunsSource -FilterDate $filterDate -FilterTopic $Topic -IncludeAllDates:([bool]$AllDates)
$remoteArtifacts = @{}
if ($IncludeRemoteBranches -or $localArtifacts.Count -eq 0) {
    $remoteArtifacts = Get-RemoteRunArtifacts -Root $RepoRoot -RunsPrefix $RunsRelativeNormalized -FilterDate $filterDate -FilterTopic $Topic -IncludeAllDates:([bool]$AllDates)
}

$artifacts = Merge-ArtifactMaps -Primary $localArtifacts -Secondary $remoteArtifacts

if ($artifacts.Count -eq 0) {
    $topicMsg = if ($Topic) { " topic=$Topic" } else { "" }
    Write-MirrorLog "No LinkedIn run artifacts found for mirror (date=$filterDate allDates=$AllDates$topicMsg)"
    exit 0
}

$copied = 0
$skipped = 0
$topicKeys = @{}

foreach ($entry in ($artifacts.GetEnumerator() | Sort-Object Name)) {
    $key = $entry.Key
    $meta = $entry.Value
    $parts = $key -split '/'
    if ($parts.Count -lt 3) { continue }

    $runDate = $parts[0]
    $topicSlug = $parts[1]
    $fileName = $parts[2]
    $topicKeys["$runDate/$topicSlug"] = $true
    $destDir = Join-Path $MirrorTarget (Join-Path $runDate $topicSlug)
    $destPath = Join-Path $destDir $fileName

    $needsCopy = $true
    if (Test-Path $destPath) {
        $destTime = (Get-Item $destPath).LastWriteTimeUtc
        if ($destTime -ge $meta.UpdatedAt) {
            $needsCopy = $false
        }
    }

    if (-not $needsCopy) {
        $skipped++
        continue
    }

    if ($DryRun) {
        Write-MirrorLog "[dry-run] would mirror $key -> $destPath"
        $copied++
        continue
    }

    New-Item -ItemType Directory -Force -Path $destDir | Out-Null

    if ($meta.SourceKind -eq "local") {
        Copy-Item -Path $meta.SourcePath -Destination $destPath -Force
    } else {
        Export-RemoteGitFile -Root $RepoRoot -Branch $meta.Branch -GitPath $meta.GitPath -DestPath $destPath
    }

    (Get-Item $destPath).LastWriteTimeUtc = $meta.UpdatedAt
    if ($fileName -like "*-infographic.png") {
        Remove-C2paFromLinkedInPng -Path $destPath
    }
    Write-MirrorLog "Mirrored $key -> $destPath"
    $copied++
}

$manifests = @()
foreach ($topicKey in ($topicKeys.Keys | Sort-Object)) {
    $parts = $topicKey -split '/'
    $manifest = Write-ReadyForPostingManifest -TargetRoot $MirrorTarget -RunDate $parts[0] -TopicSlug $parts[1] -IsDryRun:([bool]$DryRun)
    if ($null -ne $manifest) {
        $manifests += $manifest
    }
}

if ($manifests.Count -gt 0 -and -not $DryRun) {
    $summaryPath = Join-Path $LogDir "latest-mirror-runs.json"
    $manifests | ConvertTo-Json -Depth 4 | Set-Content -Path $summaryPath -Encoding UTF8
    Write-MirrorLog "Wrote mirror summary $summaryPath"
}

Write-MirrorLog "Mirror complete: copied=$copied skipped=$skipped target=$MirrorTarget template=$WindowsMirrorPathTemplate log=$LogFile"
exit 0
