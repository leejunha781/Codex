# Design Brief — engineering-rca-evidence-structure
Reference style: commissioning-gates-solution-overview

## Header
- Title (ALL CAPS): RCA EVIDENCE CHAIN: ENGINEERING GATES
- Subtitle (cyan insight line): Evidence before root cause — not narrative-and-hope
- Badge: SOLUTION OVERVIEW (top-right)

## 7 Gates (number, title, 2-3 bullets each)
1. **INTAKE REQUEST** — Field symptom log captured; Equipment ID + mission context recorded; Owner assigned
2. **SYMPTOM TIMELINE** — Time-stamped event sequence; Operator notes linked; Initial severity classified
3. **TEST EVIDENCE** — BER/SNR + terminal logs captured; Shore-gateway correlation; Measurement thresholds logged
4. **ENVIRONMENT CONTEXT** — Vessel motion + antenna state; Network routing path; Weather/obstruction noted
5. **ROOT CAUSE PROOF** — Hypothesis documented; Verification test executed; Evidence linked to claim
6. **CORRECTIVE ACTION** — Fix applied + owner; Re-test evidence captured; Change record updated
7. **CLOSURE PACK** — Engineer sign-off gate; Customer-ready DOCX/PDF export; Acceptance record filed

## Center scene
- Photorealistic subject: Naval vessel at sea twilight with SATCOM terminal/antenna visible on deck
- Overlay nodes: Field Intake → Terminal → Shore Gateway → Evidence Store → Ops Review
- TRACED LINK target on photo: Cyan solid line from overlay "Terminal" icon to actual antenna hardware on vessel
- HOLD/ESCALATE exception path: Amber dashed line + warning triangle on missing evidence link (gateway log gap)

## Left sidebar — RCA EVENTS (4 status items)
- Intake complete ✓
- Evidence linked ✓
- Root cause verified ✓
- Closure signed ✓

## Right sidebar — PYTHON AUTOMATION (4 tasks)
- Parse terminal + gateway logs
- Compare RF thresholds
- Flag missing evidence holds
- Export review-ready closure pack

## Bottom pillars (3)
1. **Traceable Baseline** — Every claim linked to captured artifact
2. **Auditable Acceptance** — Engineer sign-off before export
3. **Export Evidence** — Structured pack for customer + PLM handoff

## Footer takeaway
AI drafts faster — engineers own the evidence chain that keeps the RCA defensible.
