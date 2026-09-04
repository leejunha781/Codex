# Cursor Windows Sandbox

Cursor's Windows agent sandbox is **not** Microsoft Windows Sandbox and **not** a native Win32 jail. On Windows, Cursor runs the **Linux Landlock + seccomp sandbox inside WSL2**. Cloud Agents do not use this path.

Official references:

- [Run Modes](https://cursor.com/docs/agent/security/run-modes)
- [sandbox.json](https://cursor.com/docs/reference/sandbox)
- [Agent sandboxing (Windows = WSL2)](https://cursor.com/blog/agent-sandboxing)

## What this repo enables

Committed policy (takes effect for local desktop agents in this workspace):

| File | Role |
|------|------|
| `.cursor/sandbox.json` | Workspace read/write, shared build cache, deny-default network allowlist |
| `.cursor/permissions.json` | Auto-review allow/block instructions |
| `.cursor/rules/50-windows-sandbox.mdc` | Agent behavior when a sandbox constraint fires |
| `.cursor/skills/cursor-windows-sandbox/SKILL.md` | Diagnosis and enablement workflow |

These files cannot flip Cursor UI settings or install WSL2. You still complete the machine checklist below.

## Windows machine checklist

1. Windows 10 2004+ or Windows 11, virtualization enabled.
2. Install WSL2 and a Linux distro (Ubuntu recommended):

   ```powershell
   wsl --install -d Ubuntu
   wsl --set-default-version 2
   wsl --update
   ```

   Reboot if this is the first WSL install.

3. Inside the distro: Linux kernel **6.2+**, Landlock v3, unprivileged user namespaces on.
4. Cursor desktop **3.6+**.
5. **Settings → Agents → Approvals & Execution** = **Auto-review**.
6. **Settings → Agents → Inline Editing & Terminal → Legacy Terminal Tool** = **Off**.
7. Network mode = **sandbox.json + Defaults**.
8. Fully quit and reopen Cursor.

Do **not** use **Run Everything** if you need a sandbox. That mode has no sandbox.

## Verify

From PowerShell in this repo:

```powershell
powershell -ExecutionPolicy Bypass -File scripts/windows-sandbox-preflight.ps1
wsl -e bash scripts/windows-sandbox-preflight.sh
```

From any Linux/WSL shell:

```bash
python3 scripts/validate-sandbox-config.py
bash scripts/windows-sandbox-preflight.sh
```

Optional: Cursor **Output** panel → **Extension Host** (or Extension Host Remote) and look for `Sandbox support detected: true`.

## Network policy

`.cursor/sandbox.json` uses `networkPolicy.default: deny` plus an allowlist for GitHub, Cursor, npm, PyPI, Microsoft docs, Notion MCP, and Linear MCP.

RFC1918 and cloud-metadata addresses stay blocked by Cursor regardless of this file. Do not set `"type": "insecure_none"`.

User-only extra paths (Windows home caches, Docker config) belong in `~/.cursor/sandbox.json` using WSL paths such as `/mnt/c/Users/<you>/.docker`, not in the committed project file.

## Cloud Agents vs local Windows

| Surface | Sandbox |
|---------|---------|
| Cursor desktop on Windows | WSL2 Linux sandbox + Run Modes |
| Cursor desktop on Linux/macOS | Landlock / Seatbelt + Run Modes |
| Cursor Cloud Agent | Dedicated remote VM; **Run Modes and sandbox.json are not used** |

## Troubleshooting

| Symptom | Fix |
|---------|-----|
| No sandbox option / every command prompts | Turn **Legacy Terminal Tool** off; switch to Auto-review; install/update WSL2 |
| Preflight kernel < 6.2 | `wsl --update`, then restart the distro |
| `unshare --user` fails | Distro/AppArmor blocking user namespaces; local desktop usually ships the profile. CLI/remote only: Cursor AppArmor package from the Run Modes docs |
| Helper missing | Start Cursor so it installs `.cursor-server` / `cursorsandbox` into the distro |
| Agent retries a blocked command | Sandbox constraint fired; escalate instead of retrying. See the Windows sandbox rule |
