# Copy manual-test infographic to local E: drive
# Run in Windows PowerShell 5.1

$ErrorActionPreference = "Stop"

$Source = Join-Path $PSScriptRoot "marine-plm-ai-era-bom-validation-infographic.png"
$DestDir = "E:\Codex\2026-07-02\marine-plm-ai-era-bom-validation"
$DestFile = Join-Path $DestDir "marine-plm-ai-era-bom-validation-infographic.png"

if (-not (Test-Path $Source)) {
    throw "Source image not found: $Source"
}

if (-not (Test-Path "E:\")) {
    throw "E: drive not found. Connect or mount E: and retry."
}

New-Item -ItemType Directory -Force -Path $DestDir | Out-Null
Copy-Item -Path $Source -Destination $DestFile -Force

Write-Host "Saved:"
Write-Host "  $DestFile"
Get-Item $DestFile | Select-Object FullName, Length, LastWriteTime
