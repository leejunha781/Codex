@echo off
setlocal EnableExtensions EnableDelayedExpansion
chcp 65001 >nul
title Cursor Windows Sandbox Setup
color 0A

echo.
echo ============================================================
echo  Cursor Windows Sandbox - one-click setup / preflight
echo ============================================================
echo  This installs/updates the Codex repo branch and runs
echo  WSL2 + Landlock sandbox checks. Cloud Agents are separate.
echo ============================================================
echo.

REM Prefer a dedicated clone under the user profile (never use %%USERPROFILE%% itself as the git root).
set "REPO_URL=https://github.com/leejunha781/Codex.git"
set "BRANCH=cursor/windows-sandbox-4428"
set "REPO_DIR=%USERPROFILE%\Codex"
set "ALT_DIR=%USERPROFILE%\Codex-windows-sandbox"

where git >nul 2>&1
if errorlevel 1 (
  echo [FAIL] git.exe not found. Install Git for Windows:
  echo        https://git-scm.com/download/win
  echo.
  pause
  exit /b 1
)

where wsl >nul 2>&1
if errorlevel 1 (
  echo [WARN] wsl.exe not found. Install WSL2 after this script, then re-run.
  echo        Elevated PowerShell: wsl --install -d Ubuntu
  echo.
)

REM Pick a writable dedicated repo directory.
call :ResolveRepoDir
if errorlevel 1 (
  echo [FAIL] Could not choose a Codex clone directory.
  pause
  exit /b 1
)

echo [INFO] Using repo: !REPO_DIR!
echo [INFO] Branch:     %BRANCH%
echo.

REM Clear only THIS repo's stale lock (do not touch %%USERPROFILE%%\.git).
if exist "!REPO_DIR!\.git\index.lock" (
  echo [INFO] Removing stale lock: !REPO_DIR!\.git\index.lock
  del /f /q "!REPO_DIR!\.git\index.lock" >nul 2>&1
)

if not exist "!REPO_DIR!\.git\" (
  echo [INFO] Cloning %REPO_URL% ...
  git clone -b "%BRANCH%" "%REPO_URL%" "!REPO_DIR!"
  if errorlevel 1 (
    echo [FAIL] git clone failed.
    pause
    exit /b 1
  )
) else (
  echo [INFO] Updating existing clone ...
  pushd "!REPO_DIR!"
  git remote get-url origin >nul 2>&1
  if errorlevel 1 (
    echo [FAIL] !REPO_DIR! has .git but no origin remote.
    popd
    pause
    exit /b 1
  )
  git fetch origin "%BRANCH%"
  if errorlevel 1 (
    echo [FAIL] git fetch failed.
    popd
    pause
    exit /b 1
  )
  git checkout "%BRANCH%"
  if errorlevel 1 (
    echo [INFO] Local branch missing; creating from origin/%BRANCH% ...
    git checkout -B "%BRANCH%" "origin/%BRANCH%"
    if errorlevel 1 (
      echo [FAIL] git checkout failed.
      popd
      pause
      exit /b 1
    )
  )
  git pull --ff-only origin "%BRANCH%"
  popd
)

set "PS1=!REPO_DIR!\scripts\windows-sandbox-preflight.ps1"
set "SH_WIN=!REPO_DIR!\scripts\windows-sandbox-preflight.sh"
set "SANDBOX_JSON=!REPO_DIR!\.cursor\sandbox.json"

if not exist "!PS1!" (
  echo [FAIL] Missing: !PS1!
  echo        Branch checkout may be incomplete. Re-run this bat.
  pause
  exit /b 1
)
if not exist "!SH_WIN!" (
  echo [FAIL] Missing: !SH_WIN!
  pause
  exit /b 1
)
if not exist "!SANDBOX_JSON!" (
  echo [FAIL] Missing: !SANDBOX_JSON!
  pause
  exit /b 1
)

echo [PASS] Found preflight scripts and .cursor\sandbox.json
echo.

echo ------------------------------------------------------------
echo  Running Windows preflight
echo ------------------------------------------------------------
powershell -NoProfile -ExecutionPolicy Bypass -File "!PS1!"
set "PS_EXIT=!ERRORLEVEL!"

echo.
echo ------------------------------------------------------------
echo  Running WSL / Linux preflight
echo ------------------------------------------------------------

REM Convert C:\Users\... to /mnt/c/Users/... for bash
set "SH_WSL="
for /f "delims=" %%I in ('powershell -NoProfile -Command "$p='!SH_WIN!'; if ($p -match '^[A-Za-z]:') { $d=$p.Substring(0,1).ToLower(); $rest=$p.Substring(2) -replace '\\','/'; Write-Output ('/mnt/'+$d+$rest) } else { Write-Output $p }"') do set "SH_WSL=%%I"

if not defined SH_WSL (
  echo [FAIL] Could not convert Windows path to WSL path.
  set "SH_WSL=/mnt/c/Users/%USERNAME%/Codex/scripts/windows-sandbox-preflight.sh"
)

echo [INFO] WSL script: !SH_WSL!
where wsl >nul 2>&1
if errorlevel 1 (
  echo [WARN] Skipping WSL preflight because wsl.exe is missing.
  set "WSL_EXIT=1"
) else (
  wsl -e bash "!SH_WSL!"
  set "WSL_EXIT=!ERRORLEVEL!"
)

echo.
echo ============================================================
echo  Cursor desktop settings still required
echo ============================================================
echo  1. Settings ^> Agents ^> Approvals ^& Execution = Auto-review
echo  2. Settings ^> Agents ^> Inline Editing ^& Terminal
echo     Legacy Terminal Tool = Off
echo  3. Network mode = sandbox.json + Defaults
echo  4. Fully quit and reopen Cursor
echo.
echo  Repo used: !REPO_DIR!
echo  Open that folder in Cursor so .cursor\sandbox.json applies.
echo ============================================================
echo.

if not "!PS_EXIT!"=="0" (
  echo [RESULT] Windows preflight exit code: !PS_EXIT!
)
if not "!WSL_EXIT!"=="0" (
  echo [RESULT] WSL preflight exit code: !WSL_EXIT!
)
if "!PS_EXIT!"=="0" if "!WSL_EXIT!"=="0" (
  echo [RESULT] Preflight finished with exit code 0.
  echo          Confirm the Cursor UI settings above, then sandbox is ready.
)

echo.
pause
exit /b 0

:ResolveRepoDir
REM Use %%USERPROFILE%%\Codex when usable; otherwise alternate folder.
if not exist "%REPO_DIR%" (
  set "REPO_DIR=%REPO_DIR%"
  exit /b 0
)
if exist "%REPO_DIR%\.git\" (
  pushd "%REPO_DIR%"
  for /f "delims=" %%U in ('git remote get-url origin 2^>nul') do set "ORIGIN=%%U"
  popd
  echo !ORIGIN! | findstr /I "leejunha781/Codex" >nul
  if not errorlevel 1 (
    exit /b 0
  )
  echo [WARN] %REPO_DIR% exists but origin is not leejunha781/Codex.
  echo [WARN] Using %ALT_DIR% instead.
  set "REPO_DIR=%ALT_DIR%"
  exit /b 0
)
REM Non-git folder named Codex — do not destroy it.
echo [WARN] %REPO_DIR% exists and is not a git clone.
echo [WARN] Using %ALT_DIR% instead.
set "REPO_DIR=%ALT_DIR%"
exit /b 0
