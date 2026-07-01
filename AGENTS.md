# AGENTS.md

## Cursor Cloud specific instructions

### Repository state (important)

This repository has **no runnable application, service, or build system** on the default
`main` branch. `main` contains only:

- `README.md` (a single line: `# Codex`)
- `LICENSE` (GPLv2)
- `.gitignore` (a Unity template)

There are **no dependencies to install** (no `package.json`, `requirements.txt`,
`pyproject.toml`, `go.mod`, `Cargo.toml`, lockfiles, `Makefile`, `.devcontainer/`, or
`.cursor/environment.json`) anywhere in the tree. Consequently there is nothing to lint,
test, build, or run, and **no update/setup script is required** for this repo.

### Where the real content lives

The substantive history is on other branches (do not switch off the current branch unless
explicitly asked):

- `origin/master` / `origin/pr/plm-deck-v6-images` — snapshots of a user's AI coding-agent
  home directories (`.claude/`, `.codex/`): config, session logs, skills, automations.
- `origin/itt-work`, `origin/plm-slide-work` — Windows **PowerShell** scripts that author
  Microsoft Office documents (résumé `.docx`, PowerPoint `.pptx`) via **Office COM
  automation**.
- `origin/memory` — a few Markdown notes describing the above.

That work is **Windows-only** (requires Windows PowerShell 5.1 + Microsoft Office 2016+ via
COM). It cannot be built or run on this Linux cloud VM, and none of it is a network
service. If you are asked to work on it, expect to reason about it statically rather than
execute it here.

### Security note

Some non-`main` branches contain committed credential-like files (e.g.
`.claude/.credentials.json`). Treat any such tokens as leaked; do not use or echo them.
