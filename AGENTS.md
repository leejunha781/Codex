# Codex — AI Agent Hub

Personal AI workspace optimized for **Cursor Pro + Claude Pro + Codex Pro** orchestration.

## Identity

You assist **Joonha Lee (이준하)** — marine/PLM solution architect with 21+ years in naval shipbuilding, defence electronics, and satellite communications.

## How to Work

1. Read `MEMORY.md` for durable context before career or domain tasks.
2. Follow `.cursor/rules/` — core behavior, models, communication, resume workflow.
3. Load `.cursor/skills/` for specialized workflows (resume/JD, career, technical).
4. Execute end-to-end: investigate → implement → verify → report.

## Model Strategy

| Layer | Tool | Role |
|-------|------|------|
| IDE Agent | Cursor Pro (Codex / Claude) | Code, files, git, MCP, Cloud Agents |
| Reasoning | Claude Pro (Projects) | Long analysis, polish, strategy |
| Implementation | GPT-5.3 Codex | Heavy coding, refactors, tests |

Enable **Max Mode** for large context. See `docs/SETUP_CURSOR_PRO_CLAUDE_CODEX.md`.

## Key Commands

- `@resume-jd-workflow` — JD matching and resume tailoring
- `@cursor-windows-sandbox` — enable/diagnose Cursor Windows sandbox (WSL2)
- `@MEMORY.md` — user context hub
- `/create-rule` — add new Cursor rules from chat

## Repos & Links

- GitHub: `leejunha781/Codex`
- LinkedIn: https://www.linkedin.com/in/joonha-lee-20b518316/
- Cursor Cloud Environment: https://cursor.com/dashboard/cloud-agents/environments/r/github.com/leejunha781/codex

## Language

- Korean chat ↔ Korean replies
- Resume/cover/LinkedIn ↔ English unless specified
