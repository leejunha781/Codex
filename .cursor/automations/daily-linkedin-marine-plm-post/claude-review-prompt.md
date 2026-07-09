# Claude Pre-Post QA Review — Daily LinkedIn Marine PLM Post

Scheduled Codex automation (09:20). Runs after Cursor Cloud creates artifacts (09:00) and before mirror+post (09:35).

## Step 1 — Resolve today's run

Read the latest manifest:
- `C:\Users\namma\.cursor\automations\cache\linkedin-mirror\latest-mirror-runs.json`
- Or repo: `.cursor\automations\daily-linkedin-marine-plm-post\runs\YYYY-MM-DD\<topic-slug>\`

Required files:
- `linkedin-post.md`
- `<topic-slug>-infographic.png` (if present)

If no run found for today, check Notion Content Calendar for status **Drafting** or **QA Review** and report "no artifacts yet".

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
9. **Design grade (MANDATORY):** image must meet `professional-image-design-rule.md` — solution-overview layout, executive composition, NOT simple/generic/childish AI poster. FAIL if image looks like a basic auto-generated infographic.

**Verdict:** PASS or FAIL with specific reason. FAIL on design grade → Blocked, do not proceed to mirror+post.

## Step 3 — Image QA review

Inspect `<topic-slug>-infographic.png` (vision) against these gates:

| Gate | Check |
|------|-------|
| text-fit | All text inside intended panels/chips; no overflow |
| leader-lines | Lines terminate on correct targets; omit if ambiguous |
| overlap QA | No thumbnails over labels, clipped titles, inconsistent card spacing |
| photo-diversity | If photo thumbnails used: 8/8 unique (YARD,DASH,IFACE,REL,BOM,HANDOVER,CONTROL,VESSEL) |
| professional-grade | Reference-grade flat vector or premium engineering poster — NOT childish/cartoon/**NOT simple generic AI** |
| design-layout | Solution-overview: pipeline + central scene + sidebars + 3 cards + footer — all present |
| style match | Deep-navy background, cyan labels, flat monochrome-blue icons, 7-step pipeline layout |

**Verdict:** PASS or FAIL with specific issue.

## Step 4 — Update Notion and Linear

Read `notion-config.json` and `linear-config.json`.

Update today's Notion calendar row:
- If both post + image PASS → Status = **Ready**, Image QA = checklist, Design Grade = Pass
- If any FAIL → Status = **Blocked**, Design Grade = Fail, Notes = failure reason

Update Linear issue `LinkedIn Design QA — <topic-slug>`:
- PASS → state Approved
- FAIL → state Blocked with design grade notes

## Step 5 — Report and memory

Return concise summary:
- Topic slug
- Post QA: pass/fail
- Image QA: pass/fail (with gate details)
- Notion status updated
- Recommendation: proceed to mirror+post (09:35) or block

Append to `.codex\automations\daily-linkedin-marine-plm-post\memory.md`:

```markdown
## YYYY-MM-DD HH:MM TZ — Claude QA Review
- Topic: <topic-slug>
- Post QA: pass/fail
- Image QA: <gate results>
- Notion status: Ready | Blocked
- Recommendation: proceed | block mirror+post
```
