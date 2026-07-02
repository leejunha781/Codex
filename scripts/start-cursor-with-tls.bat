@echo off
setlocal
REM Launch Cursor with TLS environment variables applied for the current session.
REM Use this if Cursor was started before env vars were set, or Profile still fails.

set "CERT_BUNDLE=%USERPROFILE%\.cursor\certs\windows-ca-bundle.pem"
set NODE_USE_SYSTEM_CA=1
set NODE_OPTIONS=--use-system-ca
set NODE_EXTRA_CA_CERTS=%CERT_BUNDLE%
set SSL_CERT_FILE=%CERT_BUNDLE%

set "CURSOR_EXE=%LOCALAPPDATA%\Programs\cursor\Cursor.exe"
if not exist "%CURSOR_EXE%" (
  echo Cursor.exe not found at %CURSOR_EXE%
  echo Update CURSOR_EXE in this script if Cursor is installed elsewhere.
  exit /b 1
)

echo Starting Cursor with TLS env vars...
start "" "%CURSOR_EXE%"
