# Codex

Personal AI configuration hub for **Joonha Lee (이준하)** — optimized for **Cursor Pro + Claude Pro + Codex Pro**.

## What This Repo Does

- **Agent rules** (`.cursor/rules/`) — Claude-grade reasoning, model orchestration, communication, resume workflow
- **Skills** (`.cursor/skills/`) — Resume/JD matching, interview prep, PLM consulting, Windows sandbox
- **Context hub** — `MEMORY.md`, `AGENTS.md`, `CLAUDE.md`
- **MCP** — Notion + Linear via `.cursor/mcp.json`
- **Windows sandbox** — `.cursor/sandbox.json` + WSL2 Landlock; see [docs/CURSOR_WINDOWS_SANDBOX.md](docs/CURSOR_WINDOWS_SANDBOX.md)

## Quick Start

1. Clone and open in **Cursor Pro**
2. Follow [docs/SETUP_CURSOR_PRO_CLAUDE_CODEX.md](docs/SETUP_CURSOR_PRO_CLAUDE_CODEX.md) for subscription & model setup
3. Chat in Korean or English — agent follows your preferences automatically

## Key Files

| File | Purpose |
|------|---------|
| `MEMORY.md` | Career context & preferences (update here) |
| `AGENTS.md` | Project agent instructions |
| `CLAUDE.md` | Claude Code compatibility layer |
| `.cursor/rules/` | Modular Cursor rules |
| `.cursor/skills/` | Specialized workflows |
| `.cursor/sandbox.json` | Local agent sandbox network/path policy (Windows = WSL2) |
| `docs/CURSOR_WINDOWS_SANDBOX.md` | Enable Cursor Windows sandbox |

## Example Prompts

```
이 JD 분석하고 이력서 매칭 점수 줘
@resume-jd-match

FAE 면접 질문 "Tell me about a difficult integration" 답변 준비해줘
@career-interview-prep

PLM 도입 시 traceability 구조 제안해줘
@plm-systems-consulting

Windows에서 Cursor 샌드박스 켜줘
@cursor-windows-sandbox
```

### Windows sandbox one-click

Download [`Enable-Cursor-Windows-Sandbox.bat`](https://raw.githubusercontent.com/leejunha781/Codex/cursor/windows-sandbox-4428/Enable-Cursor-Windows-Sandbox.bat) and double-click. Details: [docs/CURSOR_WINDOWS_SANDBOX.md](docs/CURSOR_WINDOWS_SANDBOX.md).

## Cloud Agents

Environment: https://cursor.com/dashboard/cloud-agents/environments/r/github.com/leejunha781/codex

## License

See [LICENSE](LICENSE).
