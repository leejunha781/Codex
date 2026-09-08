---
name: ecc-plugin
description: Use the ECC Cursor adapter (skills, agents, commands) from plugins/ecc. Use when the user mentions ECC, engineering OS, TDD workflow, security-review, verification-loop, or affaan-m/ECC.
---

# ECC Plugin (Cursor adapter)

This repo adds [ECC](https://github.com/affaan-m/ECC) as a **Cursor plugin adapter**, not as a native Cursor marketplace import.

Upstream ECC is a Claude Code / Codex marketplace (`.claude-plugin/marketplace.json`). Cursor expects `.cursor-plugin/marketplace.json`, so `/plugin marketplace add https://github.com/affaan-m/ECC.git` does not load in Cursor.

## Where it lives

- Plugin catalog: `plugins/ecc/`
- Manifest: `plugins/ecc/.cursor-plugin/plugin.json`
- Repo marketplace: `.cursor-plugin/marketplace.json`
- Pin / counts: `plugins/ecc/SOURCE.json`
- Refresh: `bash scripts/install-ecc-cursor-plugin.sh`

## What to load

Use ECC **on demand** for engineering work:

- Skills under `plugins/ecc/skills/` (TDD, review, security, verification, research)
- Agents under `plugins/ecc/agents/ecc-*.md`
- Commands under `plugins/ecc/commands/`

Do **not** treat ECC as always-on context. This hub's identity stays in `MEMORY.md`, `.cursor/rules/`, and the career skills (`resume-jd-match`, `career-interview-prep`, `plm-systems-consulting`).

## What this adapter excludes

- ECC always-on Cursor rules
- ECC hook runtime
- ECC default `chrome-devtools` MCP (hub MCP remains Notion + Linear)

## Other harnesses (do not stack on this adapter)

Claude Code:

```text
/plugin marketplace add https://github.com/affaan-m/ECC
/plugin install ecc@ecc
```

Codex:

```bash
codex plugin marketplace add affaan-m/ECC
codex plugin add ecc@ecc
```
