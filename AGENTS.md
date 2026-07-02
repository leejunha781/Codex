# AGENTS.md

## Cursor Cloud specific instructions

### Node.js TLS / Cursor cloud UI errors

Symptoms:

- **Profile**: `unable to verify the first certificate`
- **Automations**: `Unable to load automations`

Both are caused by the same issue: Cursor cannot validate HTTPS to Cursor APIs (`api2.cursor.sh`, `api.cursor.com`) when SSL inspection is active.

### Windows fix (Node v24.17.0) — run all steps

**Important:** `scripts\...` paths only work inside the cloned repository folder.  
Running from `C:\Windows\System32` will fail with *"does not exist"*.

#### Option A — Standalone scripts (no repo clone; recommended)

1. Save these two files to your **Desktop** from the repo `scripts/` folder:
   - `cursor-diagnose-standalone.ps1`
   - `cursor-fix-standalone.ps1`

2. Run in PowerShell:

```powershell
powershell -ExecutionPolicy Bypass -File "$env:USERPROFILE\Desktop\cursor-fix-standalone.ps1"
```

3. Diagnose only:

```powershell
powershell -ExecutionPolicy Bypass -File "$env:USERPROFILE\Desktop\cursor-diagnose-standalone.ps1"
```

#### Option B — From cloned repo folder

```powershell
cd C:\path\to\your\codex-repo
powershell -ExecutionPolicy Bypass -File .\scripts\fix-node-tls-windows.ps1
```

Or double-click `scripts\run-fix.bat` / `scripts\run-diagnose.bat` in File Explorer.

#### Option C — Download from GitHub (one command)

```powershell
$dir = "$env:USERPROFILE\Desktop\cursor-tls-scripts"
New-Item -ItemType Directory -Force -Path $dir | Out-Null
Invoke-WebRequest -Uri "https://raw.githubusercontent.com/leejunha781/Codex/cursor/fix-node-tls-certificate-3c96/scripts/cursor-fix-standalone.ps1" -OutFile "$dir\cursor-fix.ps1"
powershell -ExecutionPolicy Bypass -File "$dir\cursor-fix.ps1"
```

**1. Run the full fix script (pick one option above).**

This script:

- Exports Windows Root/CA certificates to `%USERPROFILE%\.cursor\certs\windows-ca-bundle.pem`
- Sets `NODE_USE_SYSTEM_CA`, `NODE_OPTIONS=--use-system-ca`, `NODE_EXTRA_CA_CERTS`, `SSL_CERT_FILE`
- Patches **every** Cursor `settings.json` (user settings + profile-specific files)

```json
{
  "http.systemCertificates": true,
  "http.experimental.systemCertificatesv2": true,
  "cursor.general.disableHttp2": true
}
```

**2. In Cursor UI (required for Automations):**

- Settings → **Network** → **HTTP Compatibility Mode** → `HTTP/1.1`
- Search `systemCertificates` → enable **Http: System Certificates** and **Experimental: System Certificates V2**
- If you use **Profiles**, open **Profiles → settings.json** and confirm the same 3 keys exist
- Run **Network Diagnostic** in Settings → Network

**3. Fully quit Cursor** (tray icon included), then reopen.

**4. Retry** on Profile and Automations pages.

**5. If it still fails, run diagnostics (use full path, not `scripts\...` from System32):**

```powershell
powershell -ExecutionPolicy Bypass -File "$env:USERPROFILE\Desktop\cursor-diagnose-standalone.ps1"
```

**6. Alternative launch** (forces env vars for Cursor's process; run from repo `scripts` folder):

```bat
scripts\start-cursor-with-tls.bat
```

### If SSL inspection is active (most common remaining cause)

Ask your IT team for the **corporate root CA** `.pem` / `.crt` file, then:

```powershell
[System.Environment]::SetEnvironmentVariable("NODE_EXTRA_CA_CERTS", "C:\path\to\corp-root.pem", "User")
```

Also ask IT to **whitelist** these domains from SSL inspection:

- `*.cursor.sh`
- `*.cursorapi.com`
- `api.cursor.com`
- `cursor-cdn.com`

### Cloud agents (this repo)

`.cursor/environment.json` sets `NODE_USE_SYSTEM_CA=1`. The install script `.cursor/verify-node-tls.sh` verifies HTTPS before agents run.

Do **not** use `NODE_TLS_REJECT_UNAUTHORIZED=0` except for temporary local debugging.

### Repository state

`main` has no application runtime beyond this cloud environment configuration. Other branches hold unrelated snapshots (Unity, PowerShell, agent config exports).
