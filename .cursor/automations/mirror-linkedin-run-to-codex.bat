@echo off
setlocal
cd /d "%~dp0..\.."
echo Repo root: %CD%
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0mirror-linkedin-run-to-codex.ps1" -Latest %*
if errorlevel 1 (
  echo.
  echo Mirror failed. Try verify mode:
  echo   powershell -ExecutionPolicy Bypass -File "%~dp0mirror-linkedin-run-to-codex.ps1" -Verify
  exit /b 1
)
endlocal
