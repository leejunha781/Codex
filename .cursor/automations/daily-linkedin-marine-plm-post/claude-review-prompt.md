# Claude Pre-Post QA Review — Daily LinkedIn Marine PLM Post

Scheduled Codex automation **`daily-linkedin-claude-review`** at **09:20 daily**. Runs after Cursor Cloud creates artifacts (09:00) and **before** mirror+post (09:35).

**Hard gates:** post quality + image professional-grade. **Any FAIL → Notion Blocked → 09:35 post skipped.**

Read with (when present in this folder):
- `prompt.md` (post rules)
- `embedded-arm-topic-reference.md` (ARM Cortex-Mx/Ax posts)
- `linkedin-reference-style-commissioning-gates.md` (if present — image layout)
- `notion-config.json` / `linear-config.json` (if present)

## Step 0 — Preconditions

- Notion row for today should be **QA Review** or **Drafting** after the 09:00 run
- Required files: `linkedin-post.md`, `<topic-slug>-infographic.png`

If no artifacts for today, report `no artifacts yet` and exit without setting Notion to Ready.

## Step 1 — Resolve today's run

Check in this order:
- `C:\Users\namma\.cursor\automations\cache\linkedin-mirror\latest-mirror-runs.json`
- Repo: `.cursor\automations\daily-linkedin-marine-plm-post\runs\YYYY-MM-DD\<topic-slug>\`
- Windows mirror: `C:\Users\namma\Documents\Codex\YYYY-MM-DD\<topic-slug>\`
- Notion Content Calendar row with Status **QA Review** or **Drafting**

## Step 2 — Post quality review (HARD GATE)

Read `linkedin-post.md` and verify:

1. English-only, professional consultative tone
2. Strong hook/title with purposeful emojis (sparing, professional)
3. 3–5 short paragraphs with concrete execution context
4. 3 practical application points
5. 5–8 relevant hashtags
6. Not promotional — decision-oriented (pain, data ownership, integration risk, proof gates)
7. At least one concrete context from: BOM maturity, design-change control, supplier package, class/test records, handover evidence, CAD/ERP/MES interfaces, control-valve commissioning, RF/SATCOM validation, shipyard adoption gates, **Cortex-M/A board bring-up**, **peripheral driver validation (UART/SPI/I2C/CAN/ADC/DMA)**, **ISR/DMA latency budgets**, **bootloader/OTA acceptance**, **device-tree/BSP ownership**, or **shipboard EMC/PCB commissioning**
8. **Freelance posts (if applicable):** consultative insight only — process (intake → structure → draft → review → deliver); no price listing
9. **Embedded ARM posts (if applicable):** read `embedded-arm-topic-reference.md`. Must be engineering insight (not chip-vendor marketing). Driver/bring-up/BSP claims must mention evidence (loopback, timing budget, device-tree ownership, or acceptance gate)

**Verdict:** PASS or FAIL with specific reason.

## Step 3 — Image QA (HARD GATE)

Inspect `<topic-slug>-infographic.png` (vision):

| Gate | Check | Fail → Blocked |
|------|-------|----------------|
| professional-grade | Executive B2B infographic — not childish/cartoon/generic AI poster | YES |
| text-fit | All text inside panels/chips | YES |
| overlap QA | No clipped titles, label collisions | YES |
| leader-lines | Lines terminate on correct targets (omit if none) | YES |
| solution-overview | Pipeline or architecture layout with title, body, takeaway | YES |
| commissioning-gates-ref | If `linkedin-reference-style-commissioning-gates.md` exists: 7 gates with bullets, photorealistic center, HUD, sidebars, pillars | YES if file present |
| embedded-diagram | If ARM topic: hardware/firmware/driver diagram (bus map, boot chain, ISR/DMA, or layering) — not a generic chip marketing poster | YES if ARM topic |

**Verdict:** PASS or FAIL. Any hard-gate FAIL → Blocked.

## Step 4 — C2PA strip QA (HARD GATE when script present)

LinkedIn shows a CR "Content credentials" badge when PNG contains C2PA metadata ([LinkedIn Help](https://www.linkedin.com/help/linkedin/answer/a6282984)).

If `strip-linkedin-image-c2pa.ps1` exists, check PNG bytes for `caBX`, `c2pa`, `jumb`:

| Result | Action |
|--------|--------|
| ABSENT | PASS |
| PRESENT | FAIL → Blocked; run strip script then re-QA |

If the strip script is missing, record `c2paQA: skipped` (do not block solely on skip).

## Step 5 — Update Notion, Linear, and write verdict

**Combined verdict:**
- ALL required gates PASS → Notion **Ready**, Design Grade **Pass**, Linear **Approved** (if Linear issue exists)
- ANY required FAIL → Notion **Blocked**, Design Grade **Fail**, Linear **Blocked**, Notes = failure reason

Write `qa-verdict.json` in the run folder:

```json
{
  "date": "YYYY-MM-DD",
  "topicSlug": "<topic-slug>",
  "postQA": "pass|fail",
  "imageQA": "pass|fail",
  "embeddedArmQA": "pass|fail|n/a",
  "c2paQA": "pass|fail|skipped",
  "verdict": "ready|blocked",
  "blockedReason": "",
  "reviewedAt": "ISO-8601",
  "reviewer": "claude"
}
```

## Step 6 — Report and memory

Return concise summary:
- Topic slug
- Post QA: pass/fail
- Image QA: pass/fail
- Embedded ARM QA: pass/fail/n/a
- C2PA QA: pass/fail/skipped
- Notion status: Ready | Blocked
- Recommendation: **proceed to mirror+post (09:35)** | **BLOCK mirror+post**

Append to `.codex\automations\daily-linkedin-marine-plm-post\memory.md`:

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
