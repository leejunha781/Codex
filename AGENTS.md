# AGENTS.md

## Cursor Cloud specific instructions

### Node.js TLS certificate error (`unable to verify the first certificate`)

If the Cloud Agents **Profile** page or `npm install` / `fetch` fails with:

```text
unable to verify the first certificate; if the root CA is installed locally, try running Node.js with --use-system-ca
```

Node.js is using only its bundled CA store and does not trust a certificate that is already installed on the OS (common on corporate networks with SSL inspection).

#### Fix in this repository (cloud agents)

This repo includes `.cursor/environment.json`, which sets:

- `NODE_USE_SYSTEM_CA=1` — trust the OS certificate store
- `NODE_OPTIONS=--use-openssl-ca` — fallback on Node versions without `--use-system-ca`

The `install` script (`.cursor/verify-node-tls.sh`) checks HTTPS to `registry.npmjs.org` and `api.github.com` before the agent runs.

#### Fix on your local machine (Cursor desktop / Profile page)

Set the same variables in your **user** environment, then restart Cursor:

**Windows (PowerShell):**

```powershell
[System.Environment]::SetEnvironmentVariable("NODE_USE_SYSTEM_CA", "1", "User")
[System.Environment]::SetEnvironmentVariable("NODE_OPTIONS", "--use-openssl-ca", "User")
```

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
