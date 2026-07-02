# AGENTS.md

## Cursor Cloud specific instructions

### Node.js TLS / Cursor cloud UI errors

Symptoms:

- **Profile**: `unable to verify the first certificate`
- **Automations**: `Unable to load automations`

Both are caused by the same issue: Cursor cannot validate HTTPS to Cursor APIs (`api2.cursor.sh`, `api.cursor.com`) when SSL inspection is active.

### Windows fix (Node v24.17.0) — run all steps

**1. Run the full fix script (PowerShell as your user):**

```powershell
powershell -ExecutionPolicy Bypass -File scripts/fix-node-tls-windows.ps1
```

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

**4. If it still fails, run diagnostics and share the output:**

```powershell
powershell -ExecutionPolicy Bypass -File scripts/diagnose-node-tls-windows.ps1
```

**5. Alternative launch** (forces env vars for Cursor's process):

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
