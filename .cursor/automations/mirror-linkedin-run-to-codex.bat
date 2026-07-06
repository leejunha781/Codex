@echo off
setlocal
set INSTALL_DIR=%USERPROFILE%\.cursor\automations
set PS1=%INSTALL_DIR%\mirror-linkedin-run-to-codex.ps1
if not exist "%PS1%" (
  echo Mirror script not installed.
  echo Run once from repo:
  echo   cd C:\Users\namma\Documents\Codex
  echo   powershell -ExecutionPolicy Bypass -File .\.cursor\automations\install-linkedin-mirror.ps1
  pause
  exit /b 1
)
powershell -NoProfile -ExecutionPolicy Bypass -File "%PS1%" -Latest %*
if errorlevel 1 pause
endlocal
