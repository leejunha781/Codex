@echo off
REM Run from anywhere — always targets this repo (Codex clone).
cd /d "%~dp0"
echo Repo root: %CD%
powershell -NoProfile -ExecutionPolicy Bypass -File "%CD%\.cursor\automations\mirror-linkedin-run-to-codex.ps1" -Latest %*
if errorlevel 1 (
  echo.
  echo Mirror failed. Run verify:
  echo   powershell -ExecutionPolicy Bypass -File "%CD%\.cursor\automations\mirror-linkedin-run-to-codex.ps1" -Verify
  pause
  exit /b 1
)
pause
