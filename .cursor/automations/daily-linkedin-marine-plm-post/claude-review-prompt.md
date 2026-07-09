# Claude Pre-Post QA Review — Daily LinkedIn Marine PLM Post

Scheduled Codex automation **`daily-linkedin-claude-review`** at **09:20 daily**. Runs after Cursor Cloud creates artifacts (09:00) and **before** mirror+post (09:35).

**Hard gates:** commissioning-gates reference match + C2PA strip. **Any FAIL → Notion Blocked → 09:35 post skipped.**

## Step 0 — Preconditions

- Notion row for today should be **QA Review** (set by Cursor 09:00 run)
- Read canonical reference: `linkedin-reference-style-commissioning-gates.md`
- Config: `notion-config.json`, `linear-config.json`

If no artifacts for today, report `no artifacts yet` and exit without setting Notion to Ready.

## Step 1 — Resolve today's run

Read the latest manifest:
- `C:\Users\namma\.cursor\automations\cache\linkedin-mirror\latest-mirror-runs.json`
- Or repo: `.cursor\automations\daily-linkedin-marine-plm-post\runs\YYYY-MM-DD\<topic-slug>\`
- Or Notion Content Calendar row with Status **QA Review** or **Drafting**

Required files:
- `linkedin-post.md`
- `<topic-slug>-infographic.png`

## Step 2 — Post quality review

Read `linkedin-post.md` and verify:

1. English-only, professional consultative tone
2. Strong hook/title with purposeful emojis (sparing, professional)
3. 3–5 short paragraphs with concrete marine/shipbuilding execution context
4. 3 practical application points
5. 5–8 relevant hashtags
6. Not promotional — decision-oriented (pain, data ownership, integration risk, proof gates)
7. At least one concrete context: BOM maturity, design-change control, supplier package, class/test records, handover evidence, CAD/ERP/MES interfaces, control-valve commissioning, RF/SATCOM validation, or shipyard adoption gates
8. **Freelance posts (if applicable):** consultative insight only — show process (intake → structure → draft → review → deliver); one concrete example; no price listing; soft CTA at most; position as engineering specialist, not generic freelancer

**Verdict:** PASS or FAIL with specific reason.

## Step 3 — Image QA (commissioning-gates HARD GATE)

Inspect `<topic-slug>-infographic.png` (vision) against `linkedin-reference-style-commissioning-gates.md`:

| Gate | Check | Fail → Blocked |
|------|-------|----------------|
| commissioning-gates-ref | ALL sections present per reference doc | YES |
| gate-cards-bullets | 7 cards: icon + title + 2–3 technical bullets each | YES |
| photorealistic-center | Maritime/defense photo center + HUD overlay (not cartoon-only) | YES |
| traced-link | Cyan TRACED LINK to hardware on photo (if applicable) | YES |
| sidebars | Left events + right Python automation | YES |
| value-pillars | 3 bottom pillars + footer takeaway + SOLUTION OVERVIEW badge | YES |
| anti-simple-image | NOT generic/simple AI poster | YES |
| text-fit | All text inside panels/chips | YES |
| leader-lines | Lines terminate on correct targets | YES |
| overlap QA | No clipped titles, label collisions | YES |
| professional-grade | Executive B2B whitepaper quality | YES |

**Verdict:** PASS or FAIL. **Any commissioning-gates FAIL → Blocked.**

## Step 4 — C2PA strip QA (HARD GATE)

LinkedIn shows CR "Content credentials" badge when PNG contains C2PA metadata ([LinkedIn Help](https://www.linkedin.com/help/linkedin/answer/a6282984)).

Check PNG bytes for: `caBX`, `c2pa`, `jumb`

| Result | Action |
|--------|--------|
| ABSENT | PASS |
| PRESENT | FAIL → Blocked; run `strip-linkedin-image-c2pa.ps1` then re-QA |

## Step 5 — Update Notion, Linear, and write verdict

**Combined verdict:**
- ALL PASS (post + commissioning-gates + C2PA) → Notion **Ready**, Design Grade **Pass**, Linear **Approved**
- ANY FAIL → Notion **Blocked**, Design Grade **Fail**, Linear **Blocked**, Notes = failure reason

Write `qa-verdict.json` in run folder:

```json
{
  "date": "YYYY-MM-DD",
  "topicSlug": "<topic-slug>",
  "postQA": "pass|fail",
  "commissioningGatesQA": "pass|fail",
  "c2paQA": "pass|fail",
  "verdict": "ready|blocked",
  "blockedReason": "",
  "reviewedAt": "ISO-8601"
}
```

## Step 6 — Report and memory

Return concise summary:
- Topic slug
- Post QA: pass/fail
- Commissioning-gates QA: pass/fail
- C2PA QA: pass/fail
- Notion status: Ready | Blocked
- Recommendation: **proceed to mirror+post (09:35)** | **BLOCK mirror+post**

Append to `.codex\automations\daily-linkedin-marine-plm-post\memory.md`:

```markdown
## YYYY-MM-DD HH:MM TZ — Claude QA Review (09:20)
- Topic: <topic-slug>
- Post QA: pass/fail
- Commissioning-gates QA: pass/fail
- C2PA QA: pass/fail
- Notion status: Ready | Blocked
- Recommendation: proceed | block mirror+post
```
