# AGENTS.md

## Cursor Cloud specific instructions

### Node.js TLS certificate error (`unable to verify the first certificate`)

If the Cloud Agents **Profile** page or `npm install` / `fetch` fails with:

```text
unable to verify the first certificate; if the root CA is installed locally, try running Node.js with --use-system-ca
```

This usually means **SSL inspection** (Zscaler, Netskope, corporate proxy, Prompt Security, etc.) is re-signing HTTPS traffic. Setting only `NODE_USE_SYSTEM_CA` is often **not enough for Cursor** — Cursor also needs its own certificate settings.

### Windows fix (Node v24.17.0) — run all steps

**1. Run the full fix script (PowerShell as your user):**

```powershell
powershell -ExecutionPolicy Bypass -File scripts/fix-node-tls-windows.ps1
```

This script:

- Exports Windows Root/CA certificates to `%USERPROFILE%\.cursor\certs\windows-ca-bundle.pem`
- Sets `NODE_USE_SYSTEM_CA`, `NODE_OPTIONS=--use-system-ca`, `NODE_EXTRA_CA_CERTS`, `SSL_CERT_FILE`
- Patches `%APPDATA%\Cursor\User\settings.json`:

```json
{
  "http.systemCertificates": true,
  "http.experimental.systemCertificatesv2": true,
  "cursor.general.disableHttp2": true
}
```

**2. Fully quit Cursor** (tray icon included), then reopen.

**3. Profile page → Retry**

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

Also ask IT to **whitelist** these domains from SSL inspection (Cursor docs):

- `*.cursor.sh`
- `*.cursorapi.com`
- `cursor-cdn.com`

### Cloud agents (this repo)

`.cursor/environment.json` sets `NODE_USE_SYSTEM_CA=1`. The install script `.cursor/verify-node-tls.sh` verifies HTTPS before agents run.

Do **not** use `NODE_TLS_REJECT_UNAUTHORIZED=0` except for temporary local debugging.

### Repository state

`main` has no application runtime beyond this cloud environment configuration. Other branches hold unrelated snapshots (Unity, PowerShell, agent config exports).
