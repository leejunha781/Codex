# AGENTS.md

## Cursor Cloud specific instructions

### Node.js TLS certificate error (`unable to verify the first certificate`)

If the Cloud Agents **Profile** page or `npm install` / `fetch` fails with:

```text
unable to verify the first certificate; if the root CA is installed locally, try running Node.js with --use-system-ca
```

Node.js is using only its bundled CA store and does not trust a certificate that is already installed on the OS (common on corporate networks with SSL inspection).

#### Fix in this repository (cloud agents)

This repo includes `.cursor/environment.json`, which sets `NODE_USE_SYSTEM_CA=1` so cloud agents trust the OS certificate store. The `install` script (`.cursor/verify-node-tls.sh`) picks `--use-system-ca` on Node 22.15+ / 23.8+ / 24+ and falls back to `--use-openssl-ca` on older runtimes.

#### Fix on your local machine (Cursor desktop / Profile page)

**Windows + Node.js v24.17.0 (recommended):**

Run the helper script from PowerShell, then restart Cursor:

```powershell
powershell -ExecutionPolicy Bypass -File scripts/fix-node-tls-windows.ps1
```

Or set variables manually:

```powershell
[System.Environment]::SetEnvironmentVariable("NODE_USE_SYSTEM_CA", "1", "User")
[System.Environment]::SetEnvironmentVariable("NODE_OPTIONS", "--use-system-ca", "User")
```

Node v24.17.0 supports `--use-system-ca` natively (the flag shown in the error message).

**macOS / Linux:**

```bash
echo 'export NODE_USE_SYSTEM_CA=1' >> ~/.profile
echo 'export NODE_OPTIONS=--use-openssl-ca' >> ~/.profile
```

Restart Cursor completely after changing environment variables.

#### If the error persists (custom corporate root CA)

1. Export your organization's root CA to a `.pem` / `.crt` file.
2. Add a Cursor Cloud secret (or local env var):

   ```text
   NODE_EXTRA_CA_CERTS=C:\path\to\corporate-root-ca.pem
   ```

3. Keep `NODE_USE_SYSTEM_CA=1` enabled.

Do **not** use `NODE_TLS_REJECT_UNAUTHORIZED=0` except for temporary local debugging; it disables all TLS verification.

### Repository state

`main` has no application runtime beyond this cloud environment configuration. Other branches hold unrelated snapshots (Unity, PowerShell, agent config exports).
