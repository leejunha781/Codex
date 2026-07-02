#Requires -Version 5.1
<#
.SYNOPSIS
  Full fix for Cursor Profile TLS errors on Windows (Node v24+).

.DESCRIPTION
  Applies all known fixes when NODE_USE_SYSTEM_CA alone is not enough:
  1. Export Windows Root/CA certificates to a PEM bundle
  2. Set NODE_EXTRA_CA_CERTS, NODE_USE_SYSTEM_CA, NODE_OPTIONS
  3. Patch Cursor settings.json (systemCertificates + disableHttp2)
  4. Verify HTTPS to Cursor API endpoints

  Restart Cursor completely after running this script.
#>
Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Write-Step([string]$Message) {
  Write-Host ""
  Write-Host ">> $Message" -ForegroundColor Cyan
}

function Export-WindowsCaBundle([string]$OutputPath) {
  $stores = @(
    "Cert:\LocalMachine\Root",
    "Cert:\CurrentUser\Root",
    "Cert:\LocalMachine\CA",
    "Cert:\CurrentUser\CA"
  )

  $seen = @{}
  $lines = New-Object System.Collections.Generic.List[string]

  foreach ($store in $stores) {
    if (-not (Test-Path $store)) {
      continue
    }

    Get-ChildItem $store -ErrorAction SilentlyContinue | ForEach-Object {
      if ($null -eq $_.RawData -or $_.RawData.Length -eq 0) {
        return
      }

      $thumb = $_.Thumbprint
      if ($seen.ContainsKey($thumb)) {
        return
      }
      $seen[$thumb] = $true

      $lines.Add("-----BEGIN CERTIFICATE-----")
      $b64 = [System.Convert]::ToBase64String($_.RawData, [System.Base64FormattingOptions]::InsertLineBreaks)
      $lines.Add(($b64 -replace "`r`n", "`n"))
      $lines.Add("-----END CERTIFICATE-----")
      $lines.Add("")
    }
  }

  if ($lines.Count -eq 0) {
    throw "No certificates found in Windows certificate stores."
  }

  $parent = Split-Path -Parent $OutputPath
  if (-not (Test-Path $parent)) {
    New-Item -ItemType Directory -Path $parent -Force | Out-Null
  }

  ($lines -join "`n") | Out-File -FilePath $OutputPath -Encoding ascii -Force
  return $seen.Count
}

function Merge-CursorTlsSettings {
  $settingsPath = Join-Path $env:APPDATA "Cursor\User\settings.json"
  $settingsDir = Split-Path -Parent $settingsPath

  if (-not (Test-Path $settingsDir)) {
    New-Item -ItemType Directory -Path $settingsDir -Force | Out-Null
  }

  $desired = [ordered]@{
    "http.systemCertificates" = $true
    "http.experimental.systemCertificatesv2" = $true
    "cursor.general.disableHttp2" = $true
  }

  if (-not (Test-Path $settingsPath)) {
    $json = $desired | ConvertTo-Json -Depth 5
    $json | Out-File -FilePath $settingsPath -Encoding utf8 -Force
    return $settingsPath
  }

  $raw = Get-Content -Raw -Path $settingsPath
  if ([string]::IsNullOrWhiteSpace($raw)) {
    $raw = "{}"
  }

  $settings = $raw | ConvertFrom-Json -AsHashtable
  if ($null -eq $settings) {
    $settings = @{}
  }

  foreach ($entry in $desired.GetEnumerator()) {
    $settings[$entry.Key] = $entry.Value
  }

  ($settings | ConvertTo-Json -Depth 20) | Out-File -FilePath $settingsPath -Encoding utf8 -Force
  return $settingsPath
}

function Get-NodeTlsOptions([string]$NodeVersion) {
  $major = [int]($NodeVersion -replace '^v(\d+)\..*', '$1')
  $minor = [int]($NodeVersion -replace '^v\d+\.(\d+)\..*', '$1')

  if ($major -ge 24 -or ($major -eq 23 -and $minor -ge 8) -or ($major -eq 22 -and $minor -ge 15)) {
    return "--use-system-ca"
  }

  return "--use-openssl-ca"
}

function Test-NodeHttps([string]$Url) {
  $script = @"
fetch('$Url').then((r) => {
  if (!r.ok) throw new Error('HTTP ' + r.status);
  console.log('OK ' + r.status);
}).catch((e) => {
  console.error('FAIL ' + e.message);
  process.exit(1);
});
"@
  node -e $script
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

Write-Step "Patch Cursor settings.json"
$settingsPath = Merge-CursorTlsSettings
Write-Host "Updated $settingsPath"
Write-Host "  http.systemCertificates = true"
Write-Host "  http.experimental.systemCertificatesv2 = true"
Write-Host "  cursor.general.disableHttp2 = true"

Write-Step "Verify HTTPS"
$targets = @(
  "https://registry.npmjs.org",
  "https://api2.cursor.sh"
)
$failed = $false
foreach ($url in $targets) {
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
  Write-Host "Some checks failed. Your network may use SSL inspection (Zscaler/Netskope/etc.)." -ForegroundColor Yellow
  Write-Host "Ask IT for the corporate root CA .pem file, then run:" -ForegroundColor Yellow
  Write-Host '  [System.Environment]::SetEnvironmentVariable("NODE_EXTRA_CA_CERTS", "C:\path\to\corp-root.pem", "User")' -ForegroundColor Yellow
  Write-Host "Also ask IT to whitelist *.cursor.sh from SSL inspection if possible." -ForegroundColor Yellow
} else {
  Write-Host "All checks passed." -ForegroundColor Green
}

Write-Host ""
Write-Host "IMPORTANT: Fully quit Cursor (including tray icon), reopen it, then click Retry on Profile." -ForegroundColor Green
Write-Host "If Profile still fails, run: powershell -ExecutionPolicy Bypass -File scripts/diagnose-node-tls-windows.ps1" -ForegroundColor Green
