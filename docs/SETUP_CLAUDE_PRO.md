# Claude Pro — Setup Guide

Focused setup for **Claude Pro** (claude.ai) as the reasoning/polish layer in Joonha Lee's AI stack, alongside Cursor Pro and Codex Pro. For the full multi-tool stack, see [SETUP_CURSOR_PRO_CLAUDE_CODEX.md](SETUP_CURSOR_PRO_CLAUDE_CODEX.md).

---

## 1. Subscription

Agents cannot activate billing — this step is manual.

| Service | Plan | Sign Up | Monthly |
|---------|------|---------|---------|
| **Claude** | Pro | [claude.ai/upgrade](https://claude.ai/upgrade) | ~$20 |

Pro unlocks: higher usage limits, Projects, Artifacts, and access to the latest Claude models (Opus/Sonnet/Haiku families).

---

## 2. Create the Project

In claude.ai → **Projects** → **New Project**:

- **Name:** `Joonha Career & Codex`
- **Upload/paste as project knowledge:**
  - `MEMORY.md` (career context, preferences — update here first, re-upload after edits)
  - Canonical resume PDF/DOCX (read-only reference, never overwritten)
  - Recent JD examples relevant to active applications

### Project System Prompt

Paste into **Project Instructions**:

```
You are a career and technical advisor for Joonha Lee (이준하).
Read MEMORY.md context. Polish resume/cover English; never invent experience.
Korean for analysis; English for outward documents.
Match pipeline: Cursor drafts → you refine → Cursor saves files with new suffixes.
```

This mirrors the compatibility layer already defined in `CLAUDE.md` at the repo root, so behavior stays consistent whether Claude is invoked via Cursor or directly at claude.ai.

---

## 3. Model Selection

Within the Project, prefer:

- **Claude Sonnet 5 / Opus** — long-form reasoning, resume/cover polish, strategy discussions
- Enable **extended/long-context** behavior for large document review (JD batches, multi-page resumes)

---

## 4. When to Use Claude Pro vs Cursor

| Task | Tool |
|------|------|
| Long strategy chat, tone polish, JD/resume review | **Claude Pro (Projects)** |
| Code, git, MCP, file edits | Cursor |
| Quick IDE fix | Cursor Tab / Inline |

Claude Pro is the "editor's desk" — it never touches files or git directly. Drafts move between Cursor and Claude by copy/paste, and finished revisions are saved back into the repo from Cursor with a new filename suffix (see `MEMORY.md` naming rules — never overwrite canonical templates).

---

## 5. Verification Checklist

- [ ] Claude Pro subscription active (claude.ai/upgrade)
- [ ] Project **"Joonha Career & Codex"** created
- [ ] `MEMORY.md` uploaded as project knowledge (re-upload after edits)
- [ ] Project system prompt pasted in
- [ ] Canonical resume/cover files uploaded as read-only reference
- [ ] Confirmed Korean-in / Korean-out, English-in / English-out behavior matches `CLAUDE.md`

---

## 6. Troubleshooting

| Issue | Fix |
|-------|-----|
| Project knowledge feels stale | Re-upload `MEMORY.md` after any edit |
| Tone drifts from preferences | Re-paste the system prompt; check `MEMORY.md` Communication Preferences section |
| Unsure which tool to use | See §4 table above |

---

*Companion to [SETUP_CURSOR_PRO_CLAUDE_CODEX.md](SETUP_CURSOR_PRO_CLAUDE_CODEX.md) — maintained in `leejunha781/Codex`.*
