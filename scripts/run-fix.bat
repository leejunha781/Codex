@echo off
setlocal
cd /d "%~dp0"
echo Running Cursor TLS fix from: %CD%
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0fix-node-tls-windows.ps1"
echo.
pause
