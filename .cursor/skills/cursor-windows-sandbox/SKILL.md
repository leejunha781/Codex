---
name: cursor-windows-sandbox
description: Enable, diagnose, and configure Cursor agent sandboxing on Windows (WSL2 Landlock). Use when the user mentions Windows sandbox, Run in Sandbox, WSL2 sandbox, Legacy Terminal Tool, sandbox.json, or Auto-review on Windows.
---

# Cursor Windows Sandbox

## When to Use

- User wants Cursor's Windows sandbox / "Run in Sandbox"
- Sandbox option missing, commands always ask for approval, or WSL2/Landlock failures
- Need to edit `.cursor/sandbox.json` or `.cursor/permissions.json`

## Facts (do not invent)

- Windows sandbox = **Linux Landlock/seccomp inside WSL2**. Native Windows sandbox is not shipped.
- **Cloud Agents do not use Run Modes or this sandbox.** They run in an isolated remote VM.
- `sandbox.json` controls reach (network, extra paths). `permissions.json` steers Auto-review. They are not interchangeable.
- Legacy Terminal Tool **must be Off**. Run Everything **has no sandbox**.

## Checklist for the Windows machine

1. Install/update WSL2 (`wsl --install` then reboot if first time). Default distro version **2**.
2. Inside the distro: kernel **6.2+**, Landlock v3, `kernel.unprivileged_userns_clone=1`.
3. Cursor desktop latest (3.6+). Settings → Agents → Approvals & Execution = **Auto-review**.
4. Settings → Agents → Inline Editing & Terminal → **Legacy Terminal Tool = Off**.
5. Network mode: **sandbox.json + Defaults** (recommended).
6. Confirm this repo's `.cursor/sandbox.json` is present (do not set `type` to `insecure_none`).
7. Run from the Codex repo root (not `%USERPROFILE%`). Example for `C:\Users\namma`:

```powershell
cd C:\Users\namma\Codex
powershell -ExecutionPolicy Bypass -File C:\Users\namma\Codex\scripts\windows-sandbox-preflight.ps1
wsl -e bash /mnt/c/Users/namma/Codex/scripts/windows-sandbox-preflight.sh
```

## Policy files in this repo

| File | Job |
|------|-----|
| `.cursor/sandbox.json` | Workspace read/write, shared build cache, deny-default network allowlist for GitHub/Cursor/npm/PyPI/MCP |
| `.cursor/permissions.json` | Auto-review allow/block instructions |
| `.cursor/rules/50-windows-sandbox.mdc` | Agent behavior when a sandbox constraint fires |

Do not weaken hardcoded protections. Extra Windows-only paths belong in the user file `~/.cursor/sandbox.json` (WSL path form), not hardcoded user-home paths in the committed file.

## Diagnosis

| Symptom | Likely cause |
|---------|----------------|
| No "Run in Sandbox" / every command prompts | Legacy Terminal On, or Run Everything, or WSL2/Landlock missing |
| Preflight succeeds but UI still missing sandbox | Restart Cursor fully; check Output → Extension Host for `Sandbox support detected` |
| Network blocked inside sandbox | Domain not in allowlist; add to `.cursor/sandbox.json` or use Defaults mode |
| Same command retried in a loop | Agent ignored sandbox failure; escalate instead of retrying |

## Docs

- https://cursor.com/docs/agent/security/run-modes
- https://cursor.com/docs/reference/sandbox
- https://cursor.com/blog/agent-sandboxing
- `docs/CURSOR_WINDOWS_SANDBOX.md`
