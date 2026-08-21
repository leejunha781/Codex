param(
    [string[]]$SkillRoots = @(
        (Join-Path $env:USERPROFILE '.codex\skills'),
        (Join-Path $env:USERPROFILE '.claude\skills'),
        (Join-Path $env:USERPROFILE '.cursor\skills')
    )
)

$ErrorActionPreference = 'Stop'
$skillName = 'stm32-collaborative-development'
$required = @(
    'SKILL.md',
    'agents\openai.yaml',
    'references\multi-engine-review.md',
    'references\open-source-qualification.md'
)
$failed = $false

foreach ($root in $SkillRoots) {
    $skillPath = Join-Path $root $skillName
    foreach ($relative in $required) {
        $path = Join-Path $skillPath $relative
        if (-not (Test-Path -LiteralPath $path -PathType Leaf)) {
            Write-Error "Missing: $path" -ErrorAction Continue
            $failed = $true
        }
    }

    $skillFile = Join-Path $skillPath 'SKILL.md'
    if (Test-Path -LiteralPath $skillFile -PathType Leaf) {
        $text = Get-Content -LiteralPath $skillFile -Raw
        if ($text -notmatch '(?ms)^---\s*\r?\nname:\s*stm32-collaborative-development\s*\r?\ndescription:\s*.+?\r?\n---') {
            Write-Error "Invalid frontmatter: $skillFile" -ErrorAction Continue
            $failed = $true
        }
    }
}

if ($failed) { exit 1 }
Write-Output 'STM32 collaborative skill is present in Codex, Claude, and Cursor skill roots.'
