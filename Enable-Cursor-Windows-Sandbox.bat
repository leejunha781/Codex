@echo off
setlocal EnableExtensions
chcp 65001 >nul
title Cursor Windows Sandbox Fix
color 0A

echo.
echo ============================================================
echo  Cursor Windows Sandbox - download fix script and run
echo ============================================================
echo  Fixes wrong paths like:
echo    C:\Users\%USERNAME%\scripts\windows-sandbox-preflight.ps1
echo  by cloning into %%USERPROFILE%%\Codex (never the home git root).
echo ============================================================
echo.

set "FIX_URL=https://raw.githubusercontent.com/leejunha781/Codex/cursor/windows-sandbox-4428/scripts/Fix-Cursor-Windows-Sandbox.ps1"
set "FIX_PS1=%TEMP%\Fix-Cursor-Windows-Sandbox.ps1"

where powershell >nul 2>&1
if errorlevel 1 (
  echo [FAIL] powershell.exe not found.
  pause
  exit /b 1
)

echo [INFO] Downloading:
echo        %FIX_URL%
echo [INFO] To:
echo        %FIX_PS1%
echo.

powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "try { $u='%FIX_URL%'; $o='%FIX_PS1%'; Invoke-WebRequest -Uri $u -OutFile $o -UseBasicParsing; if (-not (Test-Path -LiteralPath $o)) { throw 'download missing' }; Write-Host '[PASS] Downloaded fix script' -ForegroundColor Green; exit 0 } catch { Write-Host ('[FAIL] Download failed: ' + $_.Exception.Message) -ForegroundColor Red; exit 1 }"
if errorlevel 1 (
  echo.
  echo [FAIL] Could not download the fix script. Check network / GitHub access.
  echo        Manual URL: %FIX_URL%
  pause
  exit /b 1
)

echo.
echo [INFO] Running fix script ...
echo.
powershell -NoProfile -ExecutionPolicy Bypass -File "%FIX_PS1%"
set "RC=%ERRORLEVEL%"

echo.
if "%RC%"=="0" (
  echo [RESULT] Fix + preflight finished OK.
) else (
  echo [RESULT] Fix script exit code: %RC%
)
echo.
pause
exit /b %RC%
