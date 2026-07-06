# Install mirror scripts to C:\Users\namma\.cursor\automations\
# Run once from the Codex repo after git pull:
#   powershell -ExecutionPolicy Bypass -File .\.cursor\automations\install-linkedin-mirror.ps1

$ErrorActionPreference = "Stop"

$RepoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..\..")).Path
$InstallDir = Join-Path $env:USERPROFILE ".cursor\automations"
$FilesToInstall = @(
    "mirror-linkedin-run-to-codex.ps1",
    "mirror-linkedin-run-to-codex.bat",
    "WINDOWS-MIRROR-GUIDE.md"
)

New-Item -ItemType Directory -Force -Path $InstallDir | Out-Null

foreach ($file in $FilesToInstall) {
    $source = Join-Path $PSScriptRoot $file
    if (-not (Test-Path -LiteralPath $source)) {
        throw "Missing installer source: $source"
    }
    Copy-Item -LiteralPath $source -Destination (Join-Path $InstallDir $file) -Force
}

$configPath = Join-Path $InstallDir "codex-repo.path"
Set-Content -LiteralPath $configPath -Value $RepoRoot -Encoding UTF8 -NoNewline

Write-Host "Installed LinkedIn mirror tools to: $InstallDir"
Write-Host "  Repo path saved: $configPath -> $RepoRoot"
Write-Host ""
Write-Host "Run from anywhere:"
Write-Host "  powershell -ExecutionPolicy Bypass -File `"$InstallDir\mirror-linkedin-run-to-codex.ps1`" -Latest"
Write-Host ""
Write-Host "Or double-click:"
Write-Host "  $InstallDir\mirror-linkedin-run-to-codex.bat"
