#Requires -Version 5.1
<#
.SYNOPSIS
  Standalone Cursor TLS fix (no other files required).

.DESCRIPTION
  Save this file anywhere (e.g. Desktop) and run:
    powershell -ExecutionPolicy Bypass -File C:\Users\YOU\Desktop\cursor-fix.ps1
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
    if (-not (Test-Path $store)) { continue }
    Get-ChildItem $store -ErrorAction SilentlyContinue | ForEach-Object {
      if ($null -eq $_.RawData -or $_.RawData.Length -eq 0) { return }
      if ($seen.ContainsKey($_.Thumbprint)) { return }
      $seen[$_.Thumbprint] = $true
      $lines.Add("-----BEGIN CERTIFICATE-----")
      $b64 = [System.Convert]::ToBase64String($_.RawData, [System.Base64FormattingOptions]::InsertLineBreaks)
      $lines.Add(($b64 -replace "`r`n", "`n"))
      $lines.Add("-----END CERTIFICATE-----")
      $lines.Add("")
    }
  }
  $parent = Split-Path -Parent $OutputPath
  if (-not (Test-Path $parent)) { New-Item -ItemType Directory -Path $parent -Force | Out-Null }
  ($lines -join "`n") | Out-File -FilePath $OutputPath -Encoding ascii -Force
  return $seen.Count
}

function Patch-SettingsFile([string]$SettingsPath) {
  $desired = @{
    "http.systemCertificates" = $true
    "http.experimental.systemCertificatesv2" = $true
    "cursor.general.disableHttp2" = $true
  }
  $dir = Split-Path -Parent $SettingsPath
  if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
  if (-not (Test-Path $SettingsPath)) {
    ($desired | ConvertTo-Json) | Out-File -FilePath $SettingsPath -Encoding utf8 -Force
    return
  }
  $settings = (Get-Content -Raw $SettingsPath) | ConvertFrom-Json -AsHashtable
  if ($null -eq $settings) { $settings = @{} }
  foreach ($k in $desired.Keys) { $settings[$k] = $desired[$k] }
  ($settings | ConvertTo-Json -Depth 20) | Out-File -FilePath $SettingsPath -Encoding utf8 -Force
}

Write-Step "Export Windows CA bundle"
$certBundle = Join-Path $env:USERPROFILE ".cursor\certs\windows-ca-bundle.pem"
$count = Export-WindowsCaBundle -OutputPath $certBundle
Write-Host "Exported $count certs -> $certBundle"

Write-Step "Set environment variables (User)"
[System.Environment]::SetEnvironmentVariable("NODE_USE_SYSTEM_CA", "1", "User")
[System.Environment]::SetEnvironmentVariable("NODE_OPTIONS", "--use-system-ca", "User")
[System.Environment]::SetEnvironmentVariable("NODE_EXTRA_CA_CERTS", $certBundle, "User")
[System.Environment]::SetEnvironmentVariable("SSL_CERT_FILE", $certBundle, "User")
$env:NODE_USE_SYSTEM_CA = "1"
$env:NODE_OPTIONS = "--use-system-ca"
$env:NODE_EXTRA_CA_CERTS = $certBundle
$env:SSL_CERT_FILE = $certBundle

Write-Step "Patch Cursor settings.json files"
$roots = @((Join-Path $env:APPDATA "Cursor"), (Join-Path $env:USERPROFILE ".cursor"))
$patched = 0
foreach ($root in $roots) {
  if (-not (Test-Path $root)) { continue }
  Get-ChildItem $root -Filter settings.json -Recurse -ErrorAction SilentlyContinue | ForEach-Object {
    Patch-SettingsFile $_.FullName
    Write-Host "Patched $($_.FullName)"
    $patched++
  }
}
if ($patched -eq 0) {
  $default = Join-Path $env:APPDATA "Cursor\User\settings.json"
  Patch-SettingsFile $default
  Write-Host "Patched $default"
}

Write-Step "Verify api.cursor.com"
node -e "fetch('https://api.cursor.com').then(r=>console.log('OK '+r.status)).catch(e=>{console.error(e.message);process.exit(1)})"

Write-Host ""
Write-Host "Done. Now:" -ForegroundColor Green
Write-Host "  1. Cursor Settings -> Network -> HTTP/1.1"
Write-Host "  2. Fully quit Cursor (tray icon too)"
Write-Host "  3. Reopen Cursor and click Retry on Profile / Automations"
