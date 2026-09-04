<#
.SYNOPSIS
  Preflight checks so Cursor's Windows agent sandbox (Linux Landlock inside WSL2) can run.

.DESCRIPTION
  Cursor does not ship a native Win32 sandbox. On Windows, sandboxed terminal
  commands run through WSL2. This script checks Windows/WSL2 prerequisites and
  prints the Cursor settings the user must confirm in the desktop app.

  Run from PowerShell (Admin not required for most checks):
    powershell -ExecutionPolicy Bypass -File scripts/windows-sandbox-preflight.ps1
#>
[CmdletBinding()]
param()

$ErrorActionPreference = "Continue"
$fail = 0
$warn = 0

function Write-Check {
    param(
        [Parameter(Mandatory = $true)][ValidateSet("PASS", "FAIL", "WARN", "INFO")][string]$Status,
        [Parameter(Mandatory = $true)][string]$Message
    )
    switch ($Status) {
        "PASS" { Write-Host "[PASS] $Message" -ForegroundColor Green }
        "FAIL" { Write-Host "[FAIL] $Message" -ForegroundColor Red; $script:fail++ }
        "WARN" { Write-Host "[WARN] $Message" -ForegroundColor Yellow; $script:warn++ }
        "INFO" { Write-Host "[INFO] $Message" -ForegroundColor Cyan }
    }
}

Write-Host "Cursor Windows sandbox preflight"
Write-Host "================================"
Write-Check -Status INFO -Message "Windows sandbox = Linux Landlock/seccomp inside WSL2 (not a native Win32 sandbox)."

# --- Windows version ---
try {
    $os = Get-CimInstance Win32_OperatingSystem
    $caption = $os.Caption
    $build = [int]$os.BuildNumber
    Write-Check -Status INFO -Message "OS: $caption (build $build)"
    if ($build -lt 19041) {
        Write-Check -Status FAIL -Message "Windows 10 2004+ / Windows 11 required for WSL2 (build 19041+)."
    }
    else {
        Write-Check -Status PASS -Message "Windows build supports WSL2."
    }
}
catch {
    Write-Check -Status WARN -Message "Could not read Windows version: $($_.Exception.Message)"
}

# --- Virtualization ---
try {
    $hv = Get-CimInstance Win32_ComputerSystem
    if ($hv.HypervisorPresent) {
        Write-Check -Status PASS -Message "HypervisorPresent = True"
    }
    else {
        Write-Check -Status WARN -Message "HypervisorPresent = False. Enable virtualization in BIOS/UEFI if WSL2 fails to start."
    }
}
catch {
    Write-Check -Status WARN -Message "Could not read HypervisorPresent."
}

# --- WSL feature / binary ---
$wslCmd = Get-Command wsl -ErrorAction SilentlyContinue
if (-not $wslCmd) {
    Write-Check -Status FAIL -Message "wsl.exe not found. In an elevated PowerShell run: wsl --install"
}
else {
    Write-Check -Status PASS -Message "wsl.exe found at $($wslCmd.Source)"

    Write-Host ""
    Write-Host "--- wsl --status ---"
    & wsl --status 2>&1 | ForEach-Object { Write-Host $_ }

    Write-Host ""
    Write-Host "--- wsl -l -v ---"
    $listOutput = & wsl -l -v 2>&1
    $listOutput | ForEach-Object { Write-Host $_ }

    # wsl.exe often emits UTF-16LE; strip NUL so -match works in Windows PowerShell 5.x
    $listText = (($listOutput | Out-String) -replace "`0", "")
    if ($listText -match "VERSION\s+2" -or $listText -match "\s2\s*$" -or $listText -match "\s2\s") {
        Write-Check -Status PASS -Message "At least one WSL2 distro is listed."
    }
    else {
        Write-Check -Status FAIL -Message "No WSL2 distro listed. Convert with: wsl --set-version <Distro> 2"
    }

    if ($listText -match "docker-desktop" -and $listText -notmatch "Ubuntu|Debian|openSUSE|Fedora") {
        Write-Check -Status WARN -Message "Only Docker Desktop WSL distros were detected. Install a real Linux distro: wsl --install -d Ubuntu"
    }
}

# --- Probe Landlock inside default WSL distro ---
if ($wslCmd) {
    $probe = @'
set -e
echo "KERNEL=$(uname -r)"
echo "UNAME=$(uname -s -m)"
if [ -r /proc/sys/kernel/unprivileged_userns_clone ]; then
  echo "USERNS=$(cat /proc/sys/kernel/unprivileged_userns_clone)"
else
  echo "USERNS=unknown"
fi
if [ -d /sys/kernel/security/lsm ]; then
  echo "LSM=$(cat /sys/kernel/security/lsm 2>/dev/null || echo missing)"
fi
python3 - <<'PY' 2>/dev/null || true
import os, platform
print("PY_KERNEL=" + platform.release())
print("PY_UID=" + str(os.getuid()))
PY
kern="$(uname -r | cut -d- -f1)"
major="$(echo "$kern" | cut -d. -f1)"
minor="$(echo "$kern" | cut -d. -f2)"
echo "KERNEL_MAJOR=$major"
echo "KERNEL_MINOR=$minor"
'@
    Write-Host ""
    Write-Host "--- WSL distro probe ---"
    $probeOut = & wsl -e bash -lc $probe 2>&1
    $probeText = (($probeOut | Out-String) -replace "`0", "")
    $probeOut | ForEach-Object { Write-Host $_ }

    if ($LASTEXITCODE -ne 0) {
        Write-Check -Status FAIL -Message "Could not run bash inside WSL. Start the distro: wsl"
    }
    else {
        if ($probeText -match "KERNEL_MAJOR=(\d+)" ) { $kMajor = [int]$Matches[1] } else { $kMajor = 0 }
        if ($probeText -match "KERNEL_MINOR=(\d+)" ) { $kMinor = [int]$Matches[1] } else { $kMinor = 0 }
        if ($kMajor -gt 6 -or ($kMajor -eq 6 -and $kMinor -ge 2)) {
            Write-Check -Status PASS -Message "WSL kernel $kMajor.$kMinor meets Landlock v3 baseline (6.2+)."
        }
        elseif ($kMajor -gt 0) {
            Write-Check -Status FAIL -Message "WSL kernel $kMajor.$kMinor is older than 6.2. Update WSL: wsl --update"
        }
        else {
            Write-Check -Status WARN -Message "Could not parse WSL kernel version from probe output."
        }

        if ($probeText -match "USERNS=1") {
            Write-Check -Status PASS -Message "Unprivileged user namespaces are enabled in WSL."
        }
        elseif ($probeText -match "USERNS=0") {
            Write-Check -Status FAIL -Message "Unprivileged user namespaces are disabled. Cursor cannot create the Linux sandbox."
        }
        else {
            Write-Check -Status WARN -Message "Could not read kernel.unprivileged_userns_clone inside WSL."
        }
    }
}

# --- Repo policy files ---
$repoRoot = Split-Path -Parent $PSScriptRoot
if (-not $repoRoot) { $repoRoot = (Get-Location).Path }
$sandboxJson = Join-Path $repoRoot ".cursor\sandbox.json"
$permissionsJson = Join-Path $repoRoot ".cursor\permissions.json"
if (Test-Path $sandboxJson) {
    Write-Check -Status PASS -Message "Found .cursor/sandbox.json"
}
else {
    Write-Check -Status FAIL -Message "Missing .cursor/sandbox.json in $repoRoot"
}
if (Test-Path $permissionsJson) {
    Write-Check -Status PASS -Message "Found .cursor/permissions.json"
}
else {
    Write-Check -Status WARN -Message "Missing .cursor/permissions.json (optional, but recommended for Auto-review)."
}

Write-Host ""
Write-Host "Cursor desktop settings you must confirm (this script cannot toggle them)"
Write-Host "-----------------------------------------------------------------------"
Write-Host "1. Settings > Agents > Approvals & Execution = Auto-review"
Write-Host "   (Allowlist + sandboxing is OK. Run Everything has NO sandbox.)"
Write-Host "2. Settings > Agents > Inline Editing & Terminal > Legacy Terminal Tool = Off"
Write-Host "3. Network mode = sandbox.json + Defaults"
Write-Host "4. Fully quit and reopen Cursor after changing those settings"
Write-Host "5. Optional: Output panel > Extension Host, look for 'Sandbox support detected: true'"
Write-Host ""
Write-Host "Then run the Linux-side checks inside WSL:"
Write-Host "  wsl -e bash '$repoRoot/scripts/windows-sandbox-preflight.sh'"
Write-Host ""

if ($fail -gt 0) {
    Write-Check -Status FAIL -Message "Preflight finished with $fail failure(s) and $warn warning(s)."
    exit 1
}
if ($warn -gt 0) {
    Write-Check -Status WARN -Message "Preflight finished with $warn warning(s) and no hard failures."
    exit 0
}
Write-Check -Status PASS -Message "Windows/WSL prerequisites look ready. Confirm the Cursor UI settings above."
exit 0
