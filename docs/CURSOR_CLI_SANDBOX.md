# Cursor CLI Sandbox — Enable & Configure

Run agent shell commands inside an OS-level sandbox (filesystem + network isolation). Two layers:

1. **On/off switch** — global CLI setting (cannot be set per-repo).
2. **Policy** — what the sandbox allows, defined in `sandbox.json` (this repo ships one at `.cursor/sandbox.json`).

---

## 1. Enable Sandbox Mode (User Action Required, one time)

Pick one:

```bash
# Persistent — writes the setting to ~/.cursor/cli-config.json
cursor-agent sandbox enable

# Per-invocation flag (values: enabled | disabled)
cursor-agent --sandbox enabled
```

To revert to plain allowlist mode:

```bash
cursor-agent sandbox disable
```

The persistent setting lives in `~/.cursor/cli-config.json` under the `sandbox` key (`sandbox.mode`, `sandbox.networkAccess`). Prefer `sandbox enable` over hand-editing — some fields are CLI-managed. Only `permissions` can be configured at project level (`.cursor/cli.json`); sandbox on/off is global.

---

## 2. What the Sandbox Enforces

| Area | Behavior |
|------|----------|
| Filesystem | Read/write inside the workspace only. `.git/config`, `.git/hooks`, `.vscode`, `.cursorignore`, and `.cursor/*.json` are always write-protected. `/tmp` writable unless `disableTmpWrite` is set. |
| Network | Blocked by default; opened per network mode. Default mode ("sandbox.json + Defaults") adds Cursor's built-in package-manager allowlist (`registry.npmjs.org`, `pypi.org`, `github.com`, …). Private IPs and cloud metadata endpoints (`169.254.169.254`) are always blocked (SSRF protection). |
| Commands | Commands needing full system access can't be sandboxed — the CLI falls back to asking for approval. |

---

## 3. This Repo's Policy — `.cursor/sandbox.json`

Per-repo policy (higher priority than `~/.cursor/sandbox.json`; team-admin and hardcoded protections cannot be weakened):

- `type: workspace_readwrite` — agent can edit files in this workspace only.
- `networkPolicy.default: deny` with an explicit allow list (GitHub, npm, PyPI).
- Deny always beats allow. Add domains to `allow` as needed (exact, `*.wildcard`, or CIDR).

Other available keys: `additionalReadwritePaths`, `additionalReadonlyPaths`, `disableTmpWrite`, `enableSharedBuildCache`.

---

## 4. Platform Notes

| Platform | Support |
|----------|---------|
| macOS | Seatbelt (`sandbox-exec`); Cursor v2.0+, no extra setup. |
| Linux | Landlock v3 + seccomp; kernel 6.2+ with unprivileged user namespaces. On AppArmor-restricted distros install `cursor-sandbox-apparmor` from `downloads.cursor.com/lab/enterprise/`. Falls back to approval prompts if unsupported. |
| Windows | Not supported — CLI uses approval/allowlist mode instead. |

Inside a sandboxed process: `CURSOR_SANDBOX` is set; on Linux use `CURSOR_ORIG_UID` / `CURSOR_ORIG_GID` for the real user (processes see UID 0 in-namespace).

---

## 5. Quick Verification

```bash
# One-off sandboxed command (no network by default)
cursor-agent sandbox run -- curl -sI https://example.com   # should fail (network denied)
cursor-agent sandbox run -- ls                             # should succeed (workspace read)
```

References: Cursor docs — [CLI parameters](https://cursor.com/docs/cli/reference/parameters), [CLI configuration](https://cursor.com/docs/cli/reference/configuration), [Run modes](https://cursor.com/docs/agent/security/run-modes), [sandbox.json reference](https://cursor.com/docs/reference/sandbox).

*Last updated: 2026-08-21 — maintained in `leejunha781/Codex`*
