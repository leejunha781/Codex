#Requires -Version 5.1
<#
.SYNOPSIS
  Standalone Cursor TLS diagnostic (no other files required).

.DESCRIPTION
  Save this file anywhere (e.g. Desktop) and run:
    powershell -ExecutionPolicy Bypass -File C:\Users\YOU\Desktop\cursor-diagnose.ps1
#>
Set-StrictMode -Version Latest
$ErrorActionPreference = "Continue"

function Write-Section([string]$Title) {
  Write-Host ""
  Write-Host "=== $Title ===" -ForegroundColor Cyan
}

function Test-NodeFetch([string]$Url) {
  $escaped = $Url -replace "'", "\'"
  node -e "fetch('$escaped').then(r=>console.log('status='+r.status)).catch(e=>{console.error('error='+e.message);process.exit(1)})" 2>&1
}

Write-Section "Working directory"
Write-Host (Get-Location)

Write-Section "Node.js"
try {
  Write-Host "Version: $(node --version)"
} catch {
  Write-Host "Node.js not found on PATH" -ForegroundColor Red
}

Write-Section "Environment variables"
foreach ($name in @("NODE_USE_SYSTEM_CA", "NODE_OPTIONS", "NODE_EXTRA_CA_CERTS", "SSL_CERT_FILE")) {
  $user = [System.Environment]::GetEnvironmentVariable($name, "User")
  Write-Host "$name = $(if ($user) { $user } else { '(not set)' })"
}

Write-Section "Cursor settings.json"
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
  Write-Host "Not found: $settingsPath" -ForegroundColor Yellow
}

Write-Section "HTTPS checks"
$targets = @(
  "https://api2.cursor.sh",
  "https://api.cursor.com",
  "https://cursor.com",
  "https://registry.npmjs.org"
)
foreach ($url in $targets) {
  Write-Host "-> $url"
  Test-NodeFetch $url
}

Write-Section "Certificate issuer (api2.cursor.sh)"
try {
  $request = [System.Net.HttpWebRequest]::Create("https://api2.cursor.sh")
  $request.Timeout = 15000
  $null = $request.GetResponse()
  $cert = $request.ServicePoint.Certificate
  if ($cert) {
    Write-Host "Issuer: $($cert.Issuer)"
    if ($cert.Issuer -notmatch "Amazon|DigiCert|Let's Encrypt|Google Trust|Cloudflare") {
      Write-Host "WARNING: Corporate SSL inspection is likely active." -ForegroundColor Yellow
    }
  }
} catch {
  Write-Host "Could not inspect: $($_.Exception.Message)" -ForegroundColor Yellow
}

Write-Host ""
Write-Host "Next: run cursor-fix-standalone.ps1 from the same folder, then restart Cursor." -ForegroundColor Green
