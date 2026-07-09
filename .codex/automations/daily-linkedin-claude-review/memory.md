# Daily LinkedIn Claude QA Review — Codex Automation Memory

Automation: `daily-linkedin-claude-review`  
Schedule: **daily 09:20** (before mirror+post at 09:35)

## Purpose

Commissioning-gates QA gate for LinkedIn post + infographic. **Any FAIL → Notion Blocked → 09:35 post skipped.**

## Reads

- `C:\Users\namma\.cursor\automations\daily-linkedin-marine-plm-post\claude-review-prompt.md`
- `linkedin-reference-style-commissioning-gates.md`
- `notion-config.json`, `linear-config.json`

## QA gates

| Gate | Fail action |
|------|-------------|
| Post quality | Blocked |
| commissioning-gates-ref | Blocked |
| C2PA strip (no caBX/c2pa) | Blocked |

## Outputs

- Notion: Ready or Blocked
- Linear: Approved or Blocked
- Run folder: `qa-verdict.json`

## Run log template

```markdown
## YYYY-MM-DD HH:MM TZ — Claude QA Review (09:20)
- Topic: <topic-slug>
- Post QA: pass/fail
- Commissioning-gates QA: pass/fail
- C2PA QA: pass/fail
- Notion status: Ready | Blocked
- Recommendation: proceed | block mirror+post
```
