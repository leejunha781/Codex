#Requires -Version 5.1
<#
.SYNOPSIS
  Diagnose Node.js / Cursor TLS certificate issues on Windows.

.DESCRIPTION
  Checks Node version, environment variables, Cursor settings, and HTTPS
  connectivity to Cursor API endpoints. Use this when Profile still shows:
  "unable to verify the first certificate".
#>
Set-StrictMode -Version Latest
$ErrorActionPreference = "Continue"

function Write-Section([string]$Title) {
  Write-Host ""
  Write-Host "=== $Title ===" -ForegroundColor Cyan
}

function Test-NodeFetch([string]$Url) {
  $script = @"
fetch('$Url').then(async (r) => {
  console.log('status=' + r.status);
}).catch((e) => {
  console.error('error=' + e.message);
  process.exit(1);
});
"@
  node -e $script 2>&1
}

Write-Section "Node.js"
try {
  Write-Host "Version: $(node --version)"
} catch {
  Write-Host "Node.js not found on PATH" -ForegroundColor Red
}

Write-Section "Environment variables"
foreach ($name in @("NODE_USE_SYSTEM_CA", "NODE_OPTIONS", "NODE_EXTRA_CA_CERTS", "SSL_CERT_FILE", "HTTP_PROXY", "HTTPS_PROXY")) {
  $user = [System.Environment]::GetEnvironmentVariable($name, "User")
  $proc = [System.Environment]::GetEnvironmentVariable($name, "Process")
  Write-Host "$name"
  Write-Host "  User    : $(if ($user) { $user } else { '(not set)' })"
  Write-Host "  Session : $(if ($proc) { $proc } else { '(not set)' })"
}

Write-Section "Cursor settings"
$settingsPath = Join-Path $env:APPDATA "Cursor\User\settings.json"
if (Test-Path $settingsPath) {
  Write-Host "Path: $settingsPath"
  $raw = Get-Content -Raw -Path $settingsPath
  foreach ($key in @(
    "http.systemCertificates",
    "http.experimental.systemCertificatesv2",
    "cursor.general.disableHttp2"
  )) {
    if ($raw -match [regex]::Escape($key)) {
      Write-Host "  $key : present"
    } else {
      Write-Host "  $key : MISSING" -ForegroundColor Yellow
    }
  }
} else {
  Write-Host "settings.json not found at $settingsPath" -ForegroundColor Yellow
}

Write-Section "HTTPS checks (Node fetch)"
$targets = @(
  "https://registry.npmjs.org",
  "https://api.github.com",
  "https://api2.cursor.sh"
)
foreach ($url in $targets) {
  Write-Host "-> $url"
  Test-NodeFetch $url
}

Write-Section "Certificate issuer hint (PowerShell)"
try {
  $request = [System.Net.HttpWebRequest]::Create("https://api2.cursor.sh")
  $request.AllowAutoRedirect = $true
  $request.Timeout = 15000
  $null = $request.GetResponse()
  $cert = $request.ServicePoint.Certificate
  if ($cert) {
    Write-Host "Subject : $($cert.Subject)"
    Write-Host "Issuer  : $($cert.Issuer)"
    if ($cert.Issuer -notmatch "Amazon|DigiCert|Let's Encrypt|Google Trust") {
      Write-Host "WARNING: SSL inspection may be active (non-public issuer)." -ForegroundColor Yellow
    }
  }
} catch {
  Write-Host "Could not inspect certificate: $($_.Exception.Message)" -ForegroundColor Yellow
}

Write-Host ""
Write-Host "If checks fail, run: powershell -ExecutionPolicy Bypass -File scripts/fix-node-tls-windows.ps1" -ForegroundColor Green
