param(
    [string]$RepositoryRoot = (Split-Path -Parent $PSScriptRoot),
    [string]$SkillsRoot = (Join-Path $env:USERPROFILE '.codex\skills'),
    [string]$SiteRoot = (Join-Path $env:USERPROFILE 'Documents\Codex\2026-08-03\sites-plugin-sites-openai-bundled-create'),
    [switch]$Push
)

$ErrorActionPreference = 'Stop'
$repo = [IO.Path]::GetFullPath($RepositoryRoot)
$skillDestination = Join-Path $repo 'skills'
$siteDestination = Join-Path $repo 'sites\stm32-driver-catalog'

if (-not (Test-Path -LiteralPath (Join-Path $repo '.git'))) {
    throw "RepositoryRoot is not a Git checkout: $repo"
}

New-Item -ItemType Directory -Force -Path $skillDestination, $siteDestination | Out-Null

Get-ChildItem -LiteralPath $SkillsRoot -Directory | Where-Object {
    Test-Path -LiteralPath (Join-Path $_.FullName 'SKILL.md')
} | ForEach-Object {
    $target = Join-Path $skillDestination $_.Name
    & robocopy $_.FullName $target /MIR /NFL /NDL /NJH /NJS /NP | Out-Null
    if ($LASTEXITCODE -gt 7) { throw "robocopy failed for skill $($_.Name): $LASTEXITCODE" }
}

$excluded = @('.git', 'node_modules', 'build', 'dist', 'work', 'outputs', '.wrangler', '_sites-preview')
$xd = $excluded | ForEach-Object { Join-Path $SiteRoot $_ }
& robocopy $SiteRoot $siteDestination /MIR /NFL /NDL /NJH /NJS /NP /XD $xd | Out-Null
if ($LASTEXITCODE -gt 7) { throw "robocopy failed for site: $LASTEXITCODE" }

$secretPattern = 'BEGIN (RSA |EC |OPENSSH )?PRIVATE KEY|ghp_[A-Za-z0-9]{20,}|sk-[A-Za-z0-9]{20,}|AKIA[0-9A-Z]{16}'
$trackedSources = @($skillDestination, $siteDestination)
$matches = Get-ChildItem -LiteralPath $trackedSources -Recurse -File | Where-Object {
    $_.Extension -notin @('.png', '.jpg', '.jpeg', '.gif', '.ico', '.woff', '.woff2', '.pdf', '.pptx', '.zip', '.webp')
} | Select-String -Pattern $secretPattern
if ($matches) { throw 'Potential secret detected; sync stopped before staging.' }

Push-Location $repo
try {
    $changed = git status --porcelain -- skills sites
    if (-not $changed) { Write-Output 'No skill or site changes.'; exit 0 }
    git add -- skills sites scripts/sync-personal-assets.ps1 scripts/install-personal-assets-sync-task.ps1
    git diff --cached --check
    git commit -m 'chore: sync personal skills and sites'
    if ($Push) { git push }
} finally {
    Pop-Location
}
