#Requires -Version 5.1
<#
.SYNOPSIS
  Fix Node.js TLS "unable to verify the first certificate" on Windows.

.DESCRIPTION
  Node.js v22.15+, v23.8+, and v24+ (including v24.17.0) support --use-system-ca.
  This script sets user-level environment variables so Node trusts certificates
  already installed in the Windows certificate store (common on corporate networks).

  Restart Cursor after running this script, then click Retry on the Profile page.
#>
Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$nodeVersion = (node --version 2>$null)
if (-not $nodeVersion) {
  Write-Error "Node.js is not installed or not on PATH."
}

Write-Host "Detected Node.js $nodeVersion"

$major = [int]($nodeVersion -replace '^v(\d+)\..*', '$1')
$minor = [int]($nodeVersion -replace '^v\d+\.(\d+)\..*', '$1')

$nodeOptions = if ($major -ge 24 -or ($major -eq 23 -and $minor -ge 8) -or ($major -eq 22 -and $minor -ge 15)) {
  "--use-system-ca"
} else {
  "--use-openssl-ca"
}

[System.Environment]::SetEnvironmentVariable("NODE_USE_SYSTEM_CA", "1", "User")
[System.Environment]::SetEnvironmentVariable("NODE_OPTIONS", $nodeOptions, "User")

Write-Host ""
Write-Host "Set user environment variables:"
Write-Host "  NODE_USE_SYSTEM_CA=1"
Write-Host "  NODE_OPTIONS=$nodeOptions"
Write-Host ""

$env:NODE_USE_SYSTEM_CA = "1"
$env:NODE_OPTIONS = $nodeOptions

Write-Host "Verifying HTTPS..."
node -e "fetch('https://registry.npmjs.org').then(r=>{if(!r.ok)throw new Error(r.status);console.log('TLS OK: registry.npmjs.org ('+r.status+')')}).catch(e=>{console.error(e.message);process.exit(1)})"

Write-Host ""
Write-Host "Done. Restart Cursor completely, then click Retry on the Profile page."
