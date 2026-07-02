#Requires -Version 5.1
<#
.SYNOPSIS
  Full fix for Cursor Profile / Automations TLS errors on Windows.

.DESCRIPTION
  Fixes:
  - Profile: unable to verify the first certificate
  - Automations: Unable to load automations

  Applies Windows CA export, Node env vars, and patches every Cursor
  settings.json (including profile-specific files).
#>
Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

. "$PSScriptRoot\Cursor-TlsCommon.ps1"

function Write-Step([string]$Message) {
  Write-Host ""
  Write-Host ">> $Message" -ForegroundColor Cyan
}

Write-Step "Detect Node.js"
$nodeVersion = node --version
Write-Host "Node.js $nodeVersion"
$nodeOptions = Get-NodeTlsOptions $nodeVersion

Write-Step "Export Windows certificate bundle"
$certDir = Join-Path $env:USERPROFILE ".cursor\certs"
$certBundle = Join-Path $certDir "windows-ca-bundle.pem"
$certCount = Export-WindowsCaBundle -OutputPath $certBundle
Write-Host "Exported $certCount certificates to $certBundle"

Write-Step "Set user environment variables"
[System.Environment]::SetEnvironmentVariable("NODE_USE_SYSTEM_CA", "1", "User")
[System.Environment]::SetEnvironmentVariable("NODE_OPTIONS", $nodeOptions, "User")
[System.Environment]::SetEnvironmentVariable("NODE_EXTRA_CA_CERTS", $certBundle, "User")
[System.Environment]::SetEnvironmentVariable("SSL_CERT_FILE", $certBundle, "User")

$env:NODE_USE_SYSTEM_CA = "1"
$env:NODE_OPTIONS = $nodeOptions
$env:NODE_EXTRA_CA_CERTS = $certBundle
$env:SSL_CERT_FILE = $certBundle

Write-Host "NODE_USE_SYSTEM_CA=1"
Write-Host "NODE_OPTIONS=$nodeOptions"
Write-Host "NODE_EXTRA_CA_CERTS=$certBundle"
Write-Host "SSL_CERT_FILE=$certBundle"

Write-Step "Patch all Cursor settings.json files (user + profiles)"
$updatedPaths = Merge-AllCursorTlsSettings
foreach ($path in $updatedPaths) {
  Write-Host "Updated $path"
}
Write-Host "  http.systemCertificates = true"
Write-Host "  http.experimental.systemCertificatesv2 = true"
Write-Host "  cursor.general.disableHttp2 = true"

Write-Step "Verify HTTPS (Profile + Automations endpoints)"
$failed = $false
foreach ($url in (Get-CursorTlsTestUrls)) {
  Write-Host "Testing $url ..."
  try {
    Test-NodeHttps $url
  } catch {
    Write-Host "  FAILED: $($_.Exception.Message)" -ForegroundColor Red
    $failed = $true
  }
}

Write-Host ""
if ($failed) {
  Write-Host "Some checks failed. SSL inspection is likely still blocking Cursor APIs." -ForegroundColor Yellow
  Write-Host "Ask IT for the corporate root CA .pem file, then run:" -ForegroundColor Yellow
  Write-Host '  [System.Environment]::SetEnvironmentVariable("NODE_EXTRA_CA_CERTS", "C:\path\to\corp-root.pem", "User")' -ForegroundColor Yellow
  Write-Host "Ask IT to whitelist *.cursor.sh, *.cursorapi.com, api.cursor.com from SSL inspection." -ForegroundColor Yellow
} else {
  Write-Host "All checks passed." -ForegroundColor Green
}

Show-CursorNetworkUiSteps

Write-Host ""
Write-Host "IMPORTANT:" -ForegroundColor Green
Write-Host "  1. Fully quit Cursor (including tray icon)"
Write-Host "  2. Reopen Cursor"
Write-Host "  3. Retry on Profile and Automations pages"
Write-Host ""
Write-Host "If Automations still fails, run:" -ForegroundColor Green
Write-Host "  powershell -ExecutionPolicy Bypass -File scripts/diagnose-node-tls-windows.ps1"
