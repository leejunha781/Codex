#Requires -Version 5.1
<#
.SYNOPSIS
  Diagnose Cursor Profile / Automations TLS issues on Windows.
#>
Set-StrictMode -Version Latest
$ErrorActionPreference = "Continue"

. "$PSScriptRoot\Cursor-TlsCommon.ps1"

function Write-Section([string]$Title) {
  Write-Host ""
  Write-Host "=== $Title ===" -ForegroundColor Cyan
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

Write-Section "Cursor settings.json files"
$settingsFiles = Get-CursorSettingsFiles
if ($settingsFiles.Count -eq 0) {
  Write-Host "No settings.json files found under Cursor profile directories." -ForegroundColor Yellow
} else {
  $requiredKeys = @(
    "http.systemCertificates",
    "http.experimental.systemCertificatesv2",
    "cursor.general.disableHttp2"
  )

  foreach ($settingsPath in $settingsFiles) {
    Write-Host ""
    Write-Host "File: $settingsPath"
    if (-not (Test-Path $settingsPath)) {
      Write-Host "  (missing)" -ForegroundColor Yellow
      continue
    }

    $raw = Get-Content -Raw -Path $settingsPath
    foreach ($key in $requiredKeys) {
      if ($raw -match [regex]::Escape($key)) {
        Write-Host "  $key : present"
      } else {
        Write-Host "  $key : MISSING" -ForegroundColor Yellow
      }
    }
  }
}

Write-Section "HTTPS checks (Profile + Automations APIs)"
foreach ($url in (Get-CursorTlsTestUrls)) {
  Write-Host "-> $url"
  try {
    Test-NodeHttps $url
  } catch {
    Write-Host "error=$($_.Exception.Message)" -ForegroundColor Red
  }
}

Write-Section "Certificate issuer hints"
foreach ($hostName in @("api2.cursor.sh", "api.cursor.com", "cursor.com")) {
  try {
    $request = [System.Net.HttpWebRequest]::Create("https://$hostName")
    $request.AllowAutoRedirect = $true
    $request.Timeout = 15000
    $null = $request.GetResponse()
    $cert = $request.ServicePoint.Certificate
    if ($cert) {
      Write-Host "$hostName"
      Write-Host "  Subject : $($cert.Subject)"
      Write-Host "  Issuer  : $($cert.Issuer)"
      if ($cert.Issuer -notmatch "Amazon|DigiCert|Let's Encrypt|Google Trust|Cloudflare") {
        Write-Host "  WARNING: SSL inspection may be active." -ForegroundColor Yellow
      }
    }
  } catch {
    Write-Host "$hostName : $($_.Exception.Message)" -ForegroundColor Yellow
  }
}

Show-CursorNetworkUiSteps

Write-Host ""
Write-Host "If checks fail, run: powershell -ExecutionPolicy Bypass -File scripts/fix-node-tls-windows.ps1" -ForegroundColor Green
