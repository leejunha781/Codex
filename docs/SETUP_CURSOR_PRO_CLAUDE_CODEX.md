# Cursor Pro + Claude Pro + Codex Pro — Setup Guide

Complete setup to run a Claude-grade AI stack inside Cursor and alongside Claude.ai / OpenAI.

---

## 1. Subscriptions (User Action Required)

Agents cannot activate billing. Configure these in your accounts:

| Service | Plan | Sign Up | Monthly |
|---------|------|---------|---------|
| **Cursor** | Pro (or Pro Plus / Ultra for heavy use) | [cursor.com/pricing](https://cursor.com/pricing) | $20+ |
| **Claude** | Pro | [claude.ai/upgrade](https://claude.ai/upgrade) | ~$20 |
| **OpenAI** | Plus (ChatGPT) or API for Codex | [openai.com/chatgpt/pricing](https://openai.com/chatgpt/pricing) | ~$20+ |

**Recommended for your usage (daily Agent + career docs):** Cursor **Pro Plus** ($60) if you hit limits often.

---

## 2. Cursor IDE Settings

### Models (Settings → Models)

Enable and pin:

- **GPT-5.3 Codex** / **GPT-5.3 Codex High** — primary coding agent
- **Claude 4.6/4.7 Opus** or **Claude Sonnet 5** — architecture & documents
- **GPT-5.4** — general agentic tasks
- **Composer 2.5** — fast Cursor-native agent (included in first-party pool)
- **Auto** — daily default to balance cost

Turn on **Max Mode** for:

- Large repos
- Resume/JD long documents
- Multi-file refactors

### Rules (Settings → Rules → User Rules)

Paste essentials or rely on this repo's `.cursor/rules/` (already configured).

### MCP (Settings → MCP)

This repo includes `.cursor/mcp.json` with Notion and Linear. In Cursor:

1. Open **Settings → MCP**
2. Confirm servers load (authenticate Notion/Linear when prompted)
3. Optional: enable **Figma** plugin from Cursor marketplace for design-to-code

### Cloud Agents

Environment: [Codex Cloud Environment](https://cursor.com/dashboard/cloud-agents/environments/r/github.com/leejunha781/codex)

- Connect repo `leejunha781/Codex`
- Use for async long tasks (JD batch, doc generation, PR fixes)

### Privacy

- **Privacy Mode** — enable if working with sensitive defence/marine data
- Note: some Claude models require data handling approval on enterprise; individual Privacy Mode is supported

---

## 3. Claude Pro (claude.ai)

### Projects Setup

Create a Claude Project: **"Joonha Career & Codex"**

Upload or paste:

- `MEMORY.md`
- Canonical resume PDF/DOCX (read-only reference)
- Recent JD examples

### System Prompt (Project Instructions)

```
You are a career and technical advisor for Joonha Lee (이준하).
Read MEMORY.md context. Polish resume/cover English; never invent experience.
Korean for analysis; English for outward documents.
Match pipeline: Cursor drafts → you refine → Cursor saves files with new suffixes.
```

### When to Use Claude vs Cursor

| Task | Tool |
|------|------|
| Code, git, MCP, files | Cursor |
| Long strategy chat, tone polish | Claude Projects |
| Quick IDE fix | Cursor Tab / Inline |

---

## 4. Codex / OpenAI

### In Cursor

Select **GPT-5.3 Codex** as Agent model for implementation work.

### ChatGPT Plus (Optional)

Use **GPT-5.3 Codex** in ChatGPT for standalone coding sessions; sync results back to this repo.

### API (Power Users)

If you exceed included Cursor API pool, add OpenAI API key in Cursor Settings → Models → API Keys for direct billing.

---

## 5. This Repo Structure

```
Codex/
├── AGENTS.md              # Main agent instructions
├── CLAUDE.md              # Claude compatibility (always on)
├── MEMORY.md              # Context hub — update here
├── chatgpt-preferences.md
├── .cursor/
│   ├── rules/             # Modular agent rules (.mdc)
│   ├── skills/            # Workflow skills
│   └── mcp.json           # MCP server config
└── docs/
    └── SETUP_CURSOR_PRO_CLAUDE_CODEX.md  # This file
```

After clone: open in Cursor — rules and skills apply automatically.

---

## 6. Recommended Daily Workflow

### Coding

1. Open repo in Cursor Pro
2. Agent model: **Auto** or **GPT-5.3 Codex**
3. `@MEMORY.md` or specific rules as needed
4. Cloud Agent for long jobs overnight

### Job Application

1. Paste JD in Cursor chat (Korean OK)
2. `@resume-jd-match` skill or rule
3. Review match score + draft bullets
4. Optional: paste draft to Claude Project for polish
5. Save new files with `_Targeted_YYYYMMDD` suffix

### Interview Prep

1. `@career-interview-prep` in chat
2. Practice 30–60 sec English answers
3. Korean debrief if needed

---

## 7. Usage & Cost Tips

- **Auto / Composer 2.5** — cheapest for routine edits (first-party pool)
- **Claude Opus / GPT-5.3 Codex Max** — use for hard tasks only
- Monitor: [cursor.com/settings](https://cursor.com/settings) → Usage
- Typical daily Agent user: $60–100/mo total across pools (per Cursor docs)

---

## 8. Verification Checklist

- [ ] Cursor Pro active; models enabled in Settings
- [ ] Claude Pro active; Project created with MEMORY context
- [ ] OpenAI/Codex access via Cursor model list
- [ ] Repo cloned; `.cursor/rules` visible in Settings → Rules
- [ ] MCP authenticated (Notion/Linear if used)
- [ ] Cloud Agent environment linked to `leejunha781/Codex`

---

## 9. Troubleshooting

| Issue | Fix |
|-------|-----|
| Rules not applied | Check `.mdc` in `.cursor/rules/`; reload window |
| Model missing | Settings → Models → show hidden models |
| Max Mode costly | Use default context for small tasks |
| MCP auth failed | Re-authenticate in Settings → MCP |
| Korean/English mix-up | `@20-communication` rule |

---

*Last updated: 2026-07-08 — maintained in `leejunha781/Codex`*
