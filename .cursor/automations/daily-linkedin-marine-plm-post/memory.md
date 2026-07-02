# Daily LinkedIn Marine PLM Post — Cursor Automation Memory

Shared with Codex automation `daily-linkedin-marine-plm-post`. Read at the start of each run. Choose a topic and angle different from prior entries when visible. When running as Cursor Cloud Agent, also read Cursor Memories (`MEMORIES.md`) for the same automation.

## Run log template (append one block per run)

```markdown
## YYYY-MM-DD HH:MM TZ
- Topic: <topic-slug>
- Angle: <one-line angle description>
- Marine execution context: <BOM maturity | design-change control | supplier package readiness | class/test records | handover evidence | CAD/ERP/MES interfaces | control-valve commissioning | RF/SATCOM validation | shipyard adoption gates | other>
- Post path: C:\Users\namma\Documents\Codex\YYYY-MM-DD\<topic-slug>\linkedin-post.md
- Image path: C:\Users\namma\Documents\Codex\YYYY-MM-DD\<topic-slug>\<topic-slug>-infographic.png
- Image QA: text-fit pass/fail | leader-lines pass/fail | overlap QA pass/fail
- LinkedIn status: composer prepared | posted | blocked
- Blocker (if any): <concise reason>
```

---

## Topic rotation log

| Date | topic-slug | angle | marine execution context | status |
|------|------------|-------|--------------------------|--------|
| 2026-06-24 | satellite-enabled-marine-digital-twin | operational discipline for data priority, PLM context, degraded-network sync | degraded-network sync | prepared |
| 2026-06-24 | satellite-enabled-marine-operations-data-loop | satellite-enabled marine operations data loop | operations data loop | blocked (LinkedIn focus) |
| 2026-07-02 | control-valve-commissioning-evidence-gate | control-valve commissioning evidence as a PLM adoption gate | control-valve commissioning evidence, class/test records, handover evidence | blocked (no LinkedIn Windows app in cloud) |

---

## Prior run history

## 2026-06-24
- Topic: Satellite-enabled marine digital twin, angle: operational discipline for data priority, PLM context, degraded-network sync.
- Created English LinkedIn post and final infographic image under C:\Users\namma\Documents\Codex\2026-06-24\satellite-enabled-marine-digital-twin\outputs.
- Prepared LinkedIn Windows app composer with text and image preview; did not click Post because current-turn confirmation is required.
- Run time: 2026-06-24 02:42:34 +09:00

## 2026-06-24 12:34 KST
- Topic: Satellite-enabled marine operations data loop.
- Created post markdown and final image under `C:\Users\namma\Documents\Codex\2026-06-24\satellite-enabled-marine-operations-data-loop`.
- Image generated as realistic maritime operations foundation with deterministic PowerShell text overlay; QA checked visually for text fit.
- LinkedIn Windows app launch/focus failed during UI automation: app stayed obscured/blank behind Cursor/Codex surfaces, so composer was not prepared and no post was submitted.

## 2026-07-02 02:39 UTC
- Topic: control-valve-commissioning-evidence-gate
- Angle: control-valve commissioning evidence as a Marine PLM adoption gate, connecting HART tests, tag/BOM maturity, supplier package readiness, class/test records, and lifecycle handover.
- Marine execution context: control-valve commissioning evidence; class/test records; handover evidence; ERP/MES readiness.
- Repo post path: `.cursor/automations/daily-linkedin-marine-plm-post/runs/2026-07-02/control-valve-commissioning-evidence-gate/linkedin-post.md`
- Repo image path: `.cursor/automations/daily-linkedin-marine-plm-post/runs/2026-07-02/control-valve-commissioning-evidence-gate/control-valve-commissioning-evidence-gate-infographic.png`
- Windows post path: `C:\Users\namma\Documents\Codex\2026-07-02\control-valve-commissioning-evidence-gate\linkedin-post.md` (not mirrored from cloud; Windows path unavailable)
- Windows image path: `C:\Users\namma\Documents\Codex\2026-07-02\control-valve-commissioning-evidence-gate\control-valve-commissioning-evidence-gate-infographic.png` (not mirrored from cloud; Windows path unavailable)
- Image QA: text-fit pass | leader-lines pass | overlap QA pass | visual inspection pass after fixing transparent-border wipe and spacing.
- LinkedIn status: blocked.
- Blocker: LinkedIn Windows app is not available in the Cursor Cloud/Linux environment; artifacts are ready for final posting.
