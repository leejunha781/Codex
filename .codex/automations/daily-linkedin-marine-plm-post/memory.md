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

## 2026-07-11 07:41 +09:00 — Claude QA Review (09:20)
- Topic: no 2026-07-11 artifacts resolved
- Post QA: not run
- Commissioning-gates QA: not run
- C2PA QA: not run
- Notion status: Blocked
- Recommendation: block mirror+post
- Notes: latest mirror manifest points to 2026-07-09 engineering-rca-evidence-structure; no 2026-07-11 run folder or mirror folder found. Notion fallback query failed because the Notion connection requires reauthentication.

## 2026-07-11 07:43:10 +09:00 — Mirror/Post Gate (09:35)
- Topic: none resolved for 2026-07-11
- Post path: none
- Image path: none
- LinkedIn status: blocked
- Blocker: Notion Content Calendar could not be read because the Notion connection requires reauthentication; hard gate was not Ready.
- Action taken: stopped before mirror and LinkedIn posting; no Notion update was possible under the expired connection.

## 2026-07-11 09:00 KST
- Topic: marine-plm-interface-acceptance-gates
- Angle: Interface acceptance gates turn Marine PLM handoffs into accountable release decisions.
- Marine execution context: CAD/ERP/MES/document-control interface readiness + BOM maturity + supplier package readiness + class/test records + handover evidence
- Post path: C:\Users\namma\Documents\Codex\2026-07-11\marine-plm-interface-acceptance-gates\linkedin-post.md
- Image path: C:\Users\namma\Documents\Codex\2026-07-11\marine-plm-interface-acceptance-gates\marine-plm-interface-acceptance-gates-infographic.png
- Image Source: Pillow Fallback after Figma/Notion/Linear MCP unavailable and Image Gen filesystem export unavailable in this run
- Image QA: commissioning-gates-ref partial pass | text-fit pass | leader-lines pass | overlap QA pass | C2PA pass (caBX/c2pa/jumb absent: True)
- LinkedIn status: ready for final posting; composer not prepared because this automation stage does not click Post without current-turn confirmation
- Blocker (if any): Figma/Notion/Linear tools not exposed; final PNG is local fallback, not Figma export
- Run time: 2026-07-11 09:00 KST

## 2026-07-11 09:05:48 +09:00 — Codex Pro backup run
- Topic: digital-thread-planning-starter
- Angle: Start digital thread planning with evidence, ownership, interface, and adoption readiness gates before platform selection.
- Marine execution context: CAD/ERP/MES/document-control interfaces + BOM maturity + class/test records + handover evidence + shipyard adoption gates
- Post path: C:\Users\namma\Documents\Codex\2026-07-11\digital-thread-planning-starter\linkedin-post.md
- Design brief path: C:\Users\namma\Documents\Codex\2026-07-11\digital-thread-planning-starter\design-brief.md
- Image path: C:\Users\namma\Documents\Codex\2026-07-11\digital-thread-planning-starter\digital-thread-planning-starter-infographic.png
- Figma file: https://www.figma.com/design/X8VBuItlzi7og1M4YJ4rvr?node-id=1-2
- Image Source: Figma MCP export
- Image QA: text-fit pass after re-export | leader-lines pass | overlap QA pass after re-export | C2PA pass (caBX/c2pa/jumb absent) | photo-diversity N/A (Figma vector/HUD layout)
- Integration status: Notion blocked by reauthentication; Linear tool unavailable in this session, so no Notion status or Linear issue update was made.
- LinkedIn status: backup artifact ready; composer not prepared and no posting attempted because final Post requires current-turn confirmation.
- Note: earlier 2026-07-11 09:00 run exists for `marine-plm-interface-acceptance-gates`; treat this `digital-thread-planning-starter` output as a Figma-exported backup unless selected for posting.
- Run time: 2026-07-11 09:05:48 +09:00
## 2026-07-11 09:23 +09:00 — Claude QA Review (09:20)
- Topic: digital-thread-planning-starter
- Post QA: pass
- Commissioning-gates QA: fail — flat schematic center; missing photorealistic maritime/defense scene and traced real-hardware link
- C2PA QA: pass
- Notion status: Blocked
- Linear status: Blocked (target state; connector unavailable)
- Recommendation: block mirror+post
## 2026-07-12 15:54 +09:00 — Claude QA Review (09:20)
- Topic: no 2026-07-12 artifacts resolved
- Post QA: not run
- Commissioning-gates QA: not run
- C2PA QA: not run
- Notion status: Blocked
- Recommendation: block mirror+post
- Notes: latest mirror manifest still points to 2026-07-09 engineering-rca-evidence-structure; no 2026-07-12 run folder found under Cursor runs or Codex mirror. Notion fallback failed because the Notion connection requires reauthentication.
