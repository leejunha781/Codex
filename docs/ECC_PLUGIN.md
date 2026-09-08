# ECC plugin for Cursor

`/add-plugin marketplace add https://github.com/affaan-m/ECC.git` cannot be applied as a native Cursor marketplace import.

ECC publishes a **Claude Code** marketplace at `.claude-plugin/marketplace.json`. Cursor marketplaces require `.cursor-plugin/marketplace.json`. The Cursor CLI command `agent plugin marketplace add` is also not available in this Cloud Agent image.

## What this repo does instead

This hub vendors a **Cursor plugin adapter** generated from ECC:

| Path | Role |
|------|------|
| `plugins/ecc/` | Cursor plugin (skills, agents, commands, scripts) |
| `plugins/ecc/.cursor-plugin/plugin.json` | Cursor plugin manifest |
| `.cursor-plugin/marketplace.json` | Repo marketplace listing `ecc` |
| `scripts/install-ecc-cursor-plugin.sh` | Refresh adapter from upstream ECC |
| `scripts/verify-ecc-plugin.sh` | Check counts and hub isolation |

Always-on ECC Cursor rules and hooks are **not** installed into `.cursor/`, so `MEMORY.md` and the career-hub rules stay in control.

## Refresh / local register

```bash
bash scripts/install-ecc-cursor-plugin.sh
bash scripts/verify-ecc-plugin.sh
```

The install script also symlinks `~/.cursor/plugins/local/ecc` → `plugins/ecc` unless you pass `--no-local`.

## Native installs (other harnesses)

Use these only in Claude Code or Codex. Do not stack them on top of this Cursor adapter in the same project.

```text
/plugin marketplace add https://github.com/affaan-m/ECC
/plugin install ecc@ecc
```

```bash
codex plugin marketplace add affaan-m/ECC
codex plugin add ecc@ecc
```

Official Cursor-only project install (invasive; not used here):

```bash
./install.sh --profile minimal --target cursor
```

That path writes ECC rules and adapter files into the host `.cursor/` directory.
