# Daily LinkedIn Claude QA Review — Codex Automation Memory

Automation: `daily-linkedin-claude-review`  
Schedule: **daily 09:20** (before mirror+post at 09:35)

## Purpose

Claude reviews today's LinkedIn post + infographic before posting. **Any FAIL → Notion Blocked → 09:35 post skipped.**

## Reads

- `C:\Users\namma\.cursor\automations\daily-linkedin-marine-plm-post\claude-review-prompt.md`
- `prompt.md`, `embedded-arm-topic-reference.md` (when present)

## QA gates

| Gate | Fail action |
|------|-------------|
| Post quality | Blocked |
| Image professional-grade | Blocked |
| Embedded ARM (if ARM topic) | Blocked |
| C2PA strip (if strip script present) | Blocked |

## Outputs

- Notion: Ready or Blocked
- Run folder: `qa-verdict.json`

## Run log template

```markdown
## YYYY-MM-DD HH:MM TZ — Claude QA Review (09:20)
- Topic: <topic-slug>
- Post QA: pass/fail
- Image QA: pass/fail
- Embedded ARM QA: pass/fail/n/a
- C2PA QA: pass/fail/skipped
- Notion status: Ready | Blocked
- Recommendation: proceed | block mirror+post
```
