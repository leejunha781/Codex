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
| 2026-07-02 | ai-era-marine-plm-evidence-ownership | AI-era evidence ownership and release governance | BOM maturity + interface handoff evidence | prepared (quality not superior) |
| 2026-07-02 | ai-era-marine-plm-evidence-ownership-v2 | premium infographic v2 with unified thumbnail/icon language and stronger hierarchy | BOM maturity + interface handoff evidence | prepared (quality verified pass) |
| 2026-07-02 | ai-era-marine-plm-evidence-ownership-v3 | diverse real-photo cuts per module (8 unique crops) | BOM maturity + interface handoff evidence | prepared (v3 quality verified pass) |

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

## 2026-07-02 03:00 UTC
- Topic: ai-era-marine-plm-evidence-ownership.
- Created sample post and image under `.cursor/automations/daily-linkedin-marine-plm-post/runs/2026-07-02/ai-era-marine-plm-evidence-ownership`.
- QA comparison result: NOT VERIFIED for superior quality versus baseline; new sample is cleaner but less visually rich.
- LinkedIn status: blocked in cloud test (no Windows app path).

## 2026-07-02 03:06 UTC
- Topic: ai-era-marine-plm-evidence-ownership-v2.
- Created refined v2 image with higher visual richness, unified thumbnail/icon language, and stronger card/typography hierarchy.
- QA comparison result: VERIFIED PASS on requested quality axes.
- LinkedIn status: ready for final posting (cloud run; no Windows app action executed).

## 2026-07-02 03:12 UTC
- Topic: ai-era-marine-plm-evidence-ownership-v3.
- Applied 8 unique real-photo crops (YARD, DASH, IFACE, REL, BOM, HANDOVER, CONTROL, VESSEL) with unified frame/icon system.
- QA comparison vs v2: VERIFIED PASS on photo diversity and premium finish.
- LinkedIn status: ready for final posting.
