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

function ConvertTo-WslPath {
    param([Parameter(Mandatory = $true)][string]$WindowsPath)
    $full = [System.IO.Path]::GetFullPath($WindowsPath)
    if ($full -match '^[A-Za-z]:') {
        $drive = $full.Substring(0, 1).ToLowerInvariant()
        $rest = $full.Substring(2) -replace '\\', '/'
        return "/mnt/$drive$rest"
    }
    return ($full -replace '\\', '/')
}

function Invoke-WslText {
    param([Parameter(Mandatory = $true)][string[]]$ArgumentList)
    $raw = & wsl @ArgumentList 2>&1
    $code = $LASTEXITCODE
    $text = (($raw | Out-String) -replace "`0", "").Trim()
    return [pscustomobject]@{ ExitCode = $code; Text = $text; Raw = $raw }
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
$wslReady = $false
if (-not $wslCmd) {
    Write-Check -Status FAIL -Message "wsl.exe not found. In an elevated PowerShell run: wsl --install -d Ubuntu"
}
else {
    Write-Check -Status PASS -Message "wsl.exe found at $($wslCmd.Source)"

    Write-Host ""
    Write-Host "--- wsl --status ---"
    $status = Invoke-WslText -ArgumentList @("--status")
    if ($status.Text) { Write-Host $status.Text }

    Write-Host ""
    Write-Host "--- wsl -l -v ---"
    $list = Invoke-WslText -ArgumentList @("-l", "-v")
    if ($list.Text) { Write-Host $list.Text }
    $listText = $list.Text

    if ($listText -match "Windows Subsystem for Linux has no installed distributions" -or
        $listText -match "no installed distributions") {
        Write-Check -Status FAIL -Message "No WSL distro installed. Elevated PowerShell: wsl --install -d Ubuntu"
    }
    elseif ($listText -match "VERSION\s+2" -or $listText -match "(?m)\s2\s*$" -or $listText -match "\s2\s") {
        Write-Check -Status PASS -Message "At least one WSL2 distro is listed."
    }
    else {
        Write-Check -Status WARN -Message "Could not confirm a WSL2 distro from wsl -l -v. If needed: wsl --set-version <Distro> 2"
    }

    if ($listText -match "docker-desktop" -and $listText -notmatch "Ubuntu|Debian|openSUSE|Fedora|kali|Alpine") {
        Write-Check -Status WARN -Message "Only Docker Desktop WSL distros were detected. Install Ubuntu: wsl --install -d Ubuntu"
    }

    # Smoke-test bash with a single argv (no multiline -lc payloads — those break under PowerShell quoting)
    Write-Host ""
    Write-Host "--- WSL bash smoke test ---"
    $smoke = Invoke-WslText -ArgumentList @("-e", "bash", "-lc", "echo WSL_BASH_OK; uname -r")
    if ($smoke.Text) { Write-Host $smoke.Text }
    if ($smoke.ExitCode -eq 0 -and $smoke.Text -match "WSL_BASH_OK") {
        Write-Check -Status PASS -Message "bash runs inside WSL."
        $wslReady = $true
    }
    else {
        Write-Check -Status FAIL -Message "Could not run bash inside WSL. Fix with: wsl --install -d Ubuntu   then   wsl --update   then open 'wsl' once."
        Write-Check -Status INFO -Message "If a distro exists but is stopped: wsl -d Ubuntu"
    }
}

# --- Probe Landlock with one-liners only (never multiline bash -lc from PowerShell) ---
if ($wslCmd -and $wslReady) {
    Write-Host ""
    Write-Host "--- WSL kernel / Landlock probe ---"

    $kernel = Invoke-WslText -ArgumentList @("-e", "uname", "-r")
    Write-Check -Status INFO -Message "WSL kernel: $($kernel.Text)"

    $kMajor = 0
    $kMinor = 0
    if ($kernel.Text -match '^(\d+)\.(\d+)') {
        $kMajor = [int]$Matches[1]
        $kMinor = [int]$Matches[2]
    }
    if ($kMajor -gt 6 -or ($kMajor -eq 6 -and $kMinor -ge 2)) {
        Write-Check -Status PASS -Message "WSL kernel $kMajor.$kMinor meets Landlock v3 baseline (6.2+)."
    }
    elseif ($kMajor -gt 0) {
        Write-Check -Status FAIL -Message "WSL kernel $kMajor.$kMinor is older than 6.2. Update WSL: wsl --update"
    }
    else {
        Write-Check -Status WARN -Message "Could not parse WSL kernel version from: $($kernel.Text)"
    }

    $userns = Invoke-WslText -ArgumentList @("-e", "bash", "-lc", "if [ -r /proc/sys/kernel/unprivileged_userns_clone ]; then cat /proc/sys/kernel/unprivileged_userns_clone; else echo unknown; fi")
    $usernsVal = ($userns.Text -split "`n" | Select-Object -Last 1).Trim()
    if ($usernsVal -eq "1") {
        Write-Check -Status PASS -Message "Unprivileged user namespaces are enabled in WSL."
    }
    elseif ($usernsVal -eq "0") {
        Write-Check -Status FAIL -Message "Unprivileged user namespaces are disabled. Cursor cannot create the Linux sandbox."
    }
    else {
        # Confirm with unshare instead of failing hard
        $unshare = Invoke-WslText -ArgumentList @("-e", "bash", "-lc", "unshare --user --map-root-user true >/dev/null 2>&1 && echo USERNS_OK || echo USERNS_FAIL")
        if ($unshare.Text -match "USERNS_OK") {
            Write-Check -Status PASS -Message "User namespaces usable via unshare (sysctl file not readable)."
        }
        else {
            Write-Check -Status WARN -Message "Could not confirm user namespaces (got '$usernsVal')."
        }
    }

    $lsm = Invoke-WslText -ArgumentList @("-e", "bash", "-lc", "if [ -r /sys/kernel/security/lsm ]; then cat /sys/kernel/security/lsm; else echo missing; fi")
    Write-Check -Status INFO -Message "LSM: $($lsm.Text)"
    if ($lsm.Text -match "landlock") {
        Write-Check -Status PASS -Message "Landlock appears in the LSM list."
    }
    else {
        Write-Check -Status WARN -Message "Landlock not listed in LSM; Cursor may still work if CONFIG_SECURITY_LANDLOCK=y."
    }
}

# --- Repo policy files ---
$repoRoot = Split-Path -Parent $PSScriptRoot
if (-not $repoRoot) { $repoRoot = (Get-Location).Path }
$sandboxJson = Join-Path $repoRoot ".cursor\sandbox.json"
$permissionsJson = Join-Path $repoRoot ".cursor\permissions.json"
$linuxPreflight = Join-Path $repoRoot "scripts\windows-sandbox-preflight.sh"
$linuxPreflightWsl = ConvertTo-WslPath -WindowsPath $linuxPreflight

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

# --- Auto-run Linux preflight with a correct /mnt path ---
if ($wslCmd -and $wslReady -and (Test-Path $linuxPreflight)) {
    Write-Host "--- Linux preflight via WSL ---"
    Write-Check -Status INFO -Message "wsl -e bash $linuxPreflightWsl"
    $linux = Invoke-WslText -ArgumentList @("-e", "bash", $linuxPreflightWsl)
    if ($linux.Text) { Write-Host $linux.Text }
    if ($linux.ExitCode -eq 0) {
        Write-Check -Status PASS -Message "Linux/WSL preflight script exited 0."
    }
    else {
        Write-Check -Status FAIL -Message "Linux/WSL preflight script exited $($linux.ExitCode)."
    }
}
else {
    Write-Host "Then run the Linux-side checks inside WSL:"
    Write-Host "  wsl -e bash $linuxPreflightWsl"
    Write-Host ""
}

if ($fail -gt 0) {
    Write-Check -Status FAIL -Message "Preflight finished with $fail failure(s) and $warn warning(s)."
    Write-Host ""
    Write-Host "If WSL bash failed, run these in an elevated PowerShell, then re-run this script:"
    Write-Host "  wsl --install -d Ubuntu"
    Write-Host "  wsl --update"
    Write-Host "  wsl -d Ubuntu"
    exit 1
}
if ($warn -gt 0) {
    Write-Check -Status WARN -Message "Preflight finished with $warn warning(s) and no hard failures."
    exit 0
}
Write-Check -Status PASS -Message "Windows/WSL prerequisites look ready. Confirm the Cursor UI settings above."
exit 0
