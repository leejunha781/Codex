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

Enable **Max Mode / Pro large context** for this workspace. See `docs/CURSOR_PRO_MAX_MODE.md` and rule `05-cursor-pro-mode`.

## Key Commands

- `@MEMORY.md` — user context hub
- `@cursor-pro-max-orchestration` — Cursor Pro Max + Claude + Codex (all substantive work)
- `@embedded-codex-workflow` — embedded firmware pipeline
- `@stm32-device-drivers` — STM32 firmware / device drivers
- `@taeha-dispense-controller` — Taeha dispense controller development
- `@resume-jd-workflow` / `@30-resume-jd-workflow` — JD matching and resume tailoring
- `/create-rule` — add new Cursor rules from chat

Skill index: `.cursor/skills/README.md` · Embedded playbook: `docs/EMBEDDED_DEV_STACK.md`

## Repos & Links

- GitHub: `leejunha781/Codex`
- LinkedIn: https://www.linkedin.com/in/joonha-lee-20b518316/
- Cursor Cloud Environment: https://cursor.com/dashboard/cloud-agents/environments/r/github.com/leejunha781/codex

## Language

- Korean chat ↔ Korean replies
- Resume/cover/LinkedIn ↔ English unless specified
