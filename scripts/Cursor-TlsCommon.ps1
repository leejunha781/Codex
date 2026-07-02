#Requires -Version 5.1
Set-StrictMode -Version Latest

function Get-CursorTlsSettings {
  return [ordered]@{
    "http.systemCertificates" = $true
    "http.experimental.systemCertificatesv2" = $true
    "cursor.general.disableHttp2" = $true
  }
}

function Get-CursorSettingsFiles {
  $roots = @(
    (Join-Path $env:APPDATA "Cursor"),
    (Join-Path $env:USERPROFILE ".cursor")
  )

  $files = New-Object System.Collections.Generic.List[string]

  foreach ($root in $roots) {
    if (-not (Test-Path $root)) {
      continue
    }

    Get-ChildItem -Path $root -Filter "settings.json" -Recurse -ErrorAction SilentlyContinue |
      ForEach-Object { $files.Add($_.FullName) }
  }

  if ($files.Count -eq 0) {
    $defaultPath = Join-Path $env:APPDATA "Cursor\User\settings.json"
    $defaultDir = Split-Path -Parent $defaultPath
    if (-not (Test-Path $defaultDir)) {
      New-Item -ItemType Directory -Path $defaultDir -Force | Out-Null
    }
    $files.Add($defaultPath)
  }

  return $files | Select-Object -Unique
}

function Merge-CursorTlsSettingsFile([string]$SettingsPath) {
  $desired = Get-CursorTlsSettings
  $settingsDir = Split-Path -Parent $SettingsPath

  if (-not (Test-Path $settingsDir)) {
    New-Item -ItemType Directory -Path $settingsDir -Force | Out-Null
  }

  if (-not (Test-Path $SettingsPath)) {
    ($desired | ConvertTo-Json -Depth 5) | Out-File -FilePath $SettingsPath -Encoding utf8 -Force
    return $SettingsPath
  }

  $raw = Get-Content -Raw -Path $SettingsPath
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

  ($settings | ConvertTo-Json -Depth 20) | Out-File -FilePath $SettingsPath -Encoding utf8 -Force
  return $SettingsPath
}

function Merge-AllCursorTlsSettings {
  $updated = @()
  foreach ($path in (Get-CursorSettingsFiles)) {
    $updated += Merge-CursorTlsSettingsFile -SettingsPath $path
  }
  return $updated
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

function Get-NodeTlsOptions([string]$NodeVersion) {
  $major = [int]($NodeVersion -replace '^v(\d+)\..*', '$1')
  $minor = [int]($NodeVersion -replace '^v\d+\.(\d+)\..*', '$1')

  if ($major -ge 24 -or ($major -eq 23 -and $minor -ge 8) -or ($major -eq 22 -and $minor -ge 15)) {
    return "--use-system-ca"
  }

  return "--use-openssl-ca"
}

function Get-CursorTlsTestUrls {
  return @(
    "https://registry.npmjs.org",
    "https://api2.cursor.sh",
    "https://api.cursor.com",
    "https://cursor.com"
  )
}

function Test-NodeHttps([string]$Url) {
  $escaped = $Url -replace "'", "\'"
  $script = @"
fetch('$escaped').then((r) => {
  if (!r.ok) throw new Error('HTTP ' + r.status);
  console.log('OK ' + r.status);
}).catch((e) => {
  console.error('FAIL ' + e.message);
  process.exit(1);
});
"@
  node -e $script
}

function Show-CursorNetworkUiSteps {
  Write-Host ""
  Write-Host "Also configure inside Cursor UI:" -ForegroundColor Yellow
  Write-Host "  1. Settings (gear) -> Network -> HTTP Compatibility Mode -> HTTP/1.1"
  Write-Host "  2. Search 'systemCertificates' and enable both system certificate options"
  Write-Host "  3. If you use Profiles, open Profiles -> settings.json and verify the same keys exist"
  Write-Host "  4. Run Network Diagnostic in Cursor Settings -> Network"
}
