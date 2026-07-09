# Daily LinkedIn Marine PLM Post — Codex Automation Memory

Shared with Cursor automation `daily-linkedin-marine-plm-post`. Read at the start of each run. Choose a topic and angle different from prior entries when visible.

**Image standard:** Every infographic must match `linkedin-reference-style-commissioning-gates.md` (LEO Commissioning Gates solution-overview). Read from `C:\Users\namma\.cursor\automations\daily-linkedin-marine-plm-post\`.

## Run log template (append one block per run)

```markdown
## YYYY-MM-DD HH:MM TZ
- Topic: <topic-slug>
- Angle: <one-line angle description>
- Marine execution context: <BOM maturity | design-change control | supplier package readiness | class/test records | handover evidence | CAD/ERP/MES interfaces | control-valve commissioning | RF/SATCOM validation | shipyard adoption gates | other>
- Post path: C:\Users\namma\Documents\Codex\YYYY-MM-DD\<topic-slug>\linkedin-post.md
- Image path: C:\Users\namma\Documents\Codex\YYYY-MM-DD\<topic-slug>\<topic-slug>-infographic.png
- Image QA: commissioning-gates-ref pass/fail | text-fit pass/fail | leader-lines pass/fail | overlap QA pass/fail
- LinkedIn status: composer prepared | posted | blocked
- Blocker (if any): <concise reason>
```

---

## Local mirror + auto-post workflow (2026-07-06)

**Fixed Windows mirror folder:** `C:\Users\namma\Documents\Codex\YYYY-MM-DD\<topic-slug>\`
Codex automation `daily-linkedin-mirror-and-post` (09:35) mirrors repo runs then auto-posts via LinkedIn Windows app.

---

## Figma + Notion + Claude integration (2026-07-09)

**Notion Content Calendar:** https://app.notion.com/p/7ee8f488584e4ca290f3fcbfa3ea1314
**Canonical image reference:** `linkedin-reference-style-commissioning-gates.md`

Daily pipeline:

| Time | Automation | Action |
|------|-----------|--------|
| 09:00 | Cursor Cloud `daily-linkedin-marine-plm-post` | Notion topic → post + commissioning-gates image → repo |
| 09:20 | Codex `daily-linkedin-claude-review` | QA vs commissioning-gates reference |
| 09:35 | Codex `daily-linkedin-mirror-and-post` | Mirror → LinkedIn post |

Figma account: View seat (leejunha781@gmail.com). Figma MCP primary; Image Gen when blocked.

Linear project: https://linear.app/joonha-lee/project/linkedin-content-pipeline-bf77c4a8e3b6

---

## Topic rotation log

| Date | topic-slug | angle | marine execution context | status |
|------|------------|-------|--------------------------|--------|
| 2026-06-24 | satellite-enabled-marine-digital-twin | operational discipline for data priority, PLM context, degraded-network sync | degraded-network sync | prepared |
| 2026-06-24 | satellite-enabled-marine-operations-data-loop | satellite-enabled marine operations data loop | operations data loop | blocked (LinkedIn focus) |
| 2026-07-02 | ai-era-marine-plm-evidence-ownership | AI-era evidence ownership and release governance | BOM maturity + interface handoff evidence | prepared |
| 2026-07-02 | satcom-fallback-evidence-management | SATCOM fallback as engineering evidence | RF/SATCOM validation | prepared |
| 2026-07-02 | marine-plm-bom-validation-gates-test | BOM validation gates before shipyard release | BOM maturity + class/test records | prepared |
| 2026-07-09 | engineering-rca-evidence-structure | RCA reports fail without traceable evidence chains | SATCOM link-loss RCA + naval acceptance | ready for final posting (v3 commissioning-gates image) |
