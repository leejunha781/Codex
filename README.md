# Codex

Personal AI configuration hub for **Joonha Lee (이준하)** — optimized for **Cursor Pro + Claude Pro + Codex Pro**.

## What This Repo Does

- **Agent rules** (`.cursor/rules/`) — Claude-grade reasoning, model orchestration, communication, resume workflow
- **Skills** (`.cursor/skills/`) — Resume/JD matching, interview prep, PLM consulting
- **Context hub** — `MEMORY.md`, `AGENTS.md`, `CLAUDE.md`
- **MCP** — Notion + Linear via `.cursor/mcp.json`

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

## Example Prompts

```
이 JD 분석하고 이력서 매칭 점수 줘
@resume-jd-match

FAE 면접 질문 "Tell me about a difficult integration" 답변 준비해줘
@career-interview-prep

PLM 도입 시 traceability 구조 제안해줘
@plm-systems-consulting
```

## Cloud Agents

Environment: https://cursor.com/dashboard/cloud-agents/environments/r/github.com/leejunha781/codex

## License

See [LICENSE](LICENSE).
