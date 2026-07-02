@echo off
setlocal
cd /d "%~dp0"
echo Running Cursor TLS diagnostics from: %CD%
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0diagnose-node-tls-windows.ps1"
echo.
pause
