# Cursor Pro Mode / Max Context — How to Turn On

Agents **cannot** flip the IDE toggle. Do this once in Cursor Desktop.

## A. Usage-based Cursor Pro (most common)

1. Open **Agent** (`Ctrl+I`) or Chat.
2. Click the **model picker** (top of the panel).
3. Hover the model → **Edit**.
4. Set **Context** to the largest option (e.g. **1M**).
5. Optionally set **Thinking** / **Fast** as needed.
6. Choice sticks for that model.

Cycle models: `Ctrl+/`.

## B. Legacy request-based plan

1. Open Agent/Chat → **model picker**.
2. Toggle **Max Mode** **ON** (persists across chats).

## C. Enable models globally

1. **Cursor Settings** → `Ctrl+Shift+J` (or gear → Cursor Settings).
2. Open **Models**.
3. Enable: **Claude Opus/Sonnet**, **GPT-5.3 Codex**, **GPT-5.4** (use **Show hidden models** if needed).
4. Default Agent Model: **Auto** for daily; switch to Codex/Claude for heavy work.

## This Repo Already “Sets” Pro Mode Policy

| Artifact | Effect |
|----------|--------|
| `.cursor/rules/05-cursor-pro-mode.mdc` | Always-on Pro Mode policy |
| `.cursor/rules/10-model-orchestration.mdc` | Claude + Codex routing |
| `.cursor/skills/cursor-pro-max-orchestration/` | Session checklist + skill router |
| `.cursor/agents/*.md` | High-context custom agents for embedded |

After enabling context in the picker, start with:

```text
@cursor-pro-max-orchestration
@MEMORY.md
```

## Cost note

Larger context uses more tokens. Use Auto/small context for one-line fixes; keep Pro/Max for firmware, multi-file, JD, RCA.
