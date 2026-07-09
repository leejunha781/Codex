# PLM Concept Deck — Sample Outline (10 slides)

**Title:** Digital Thread Readiness for Shipbuilding — Concept Deck  
**Audience:** Engineering director, IT/PLM sponsor, program manager  
**Prepared by:** Joonha Lee — PLM & Digital Thread Planning Support  
**Date:** 2026-07-09  
**Note:** Portfolio outline; full visual deck built via `plm-slide-builder` skill and PowerPoint COM.

---

## Slide 1 — Title

**Digital Thread Readiness for Shipbuilding**  
From design release to production and class evidence  
Prepared for: [Anonymized shipyard]  
Joonha Lee | PLM & System Integration

---

## Slide 2 — Executive thesis

- Shipyards lose schedule and rework cost when engineering data, change impact, and evidence are fragmented.
- A **governed digital thread** connects authoring tools, PLM objects, change control, and verification artifacts.
- Goal: one traceable path from requirement → design → release → manufacturing → evidence.

---

## Slide 3 — Current state (typical pain)

| Area | Symptom | Business impact |
|------|---------|-----------------|
| BOM / configuration | Multiple Excel sources | Wrong material to block |
| Change control | Email + offline redlines | Late impact discovery |
| Drawings | CAD vault vs PLM mismatch | Rework at production |
| Evidence | Scattered PDFs | Class / audit risk |

---

## Slide 4 — Target operating model

- **Authoring:** AVEVA E3D/Marine/Draw on Windows workstations  
- **Control plane:** API-based PLM (objects, structures, ECR/ECO, approvals) on Linux cluster  
- **Bridge:** Local Agent / Plugin callbacks — not direct CAD-on-Linux  
- **Cloud:** CONNECT for intelligence and optional AI assistant (governed)

---

## Slide 5 — Digital thread data objects

```
Requirement → Design object → E-BOM/MTO → Baseline → ECR → ECO → M-BOM → Evidence pack
```

Key gates: AI/rule gate at promote steps; effectivity (hull/block/option/date); approval workflow.

---

## Slide 6 — Vendor landscape (summary)

| Capability | AVEVA | Siemens TC | Dassault 3DX | Hexagon |
|------------|-------|------------|--------------|---------|
| Marine 3D authoring | Strong | Partner | Partner | Niche |
| Unified engineering | Strong | Strong | Strong | Moderate |
| Open API / custom PLM layer | Flexible | Mature | Mature | Varies |
| Shipyard references | High (target yard) | High | High | Selective |

*Full benchmark available as separate deliverable.*

---

## Slide 7 — Phased roadmap (MVP)

| Phase | Duration | Outcome |
|-------|----------|---------|
| Phase 0 — Discovery | 4 wks | Process map, data objects, gap list |
| Phase 1 — MVP thread | 12 wks | One hull class: req → ECO → evidence |
| Phase 2 — Scale | 6 mo | Multi-discipline, ERP/MES touchpoints |
| Phase 3 — Optimize | Ongoing | KPI dashboard, AI-assisted checks |

---

## Slide 8 — KPI model

| KPI | Definition | Gate |
|-----|------------|------|
| Release readiness | % objects with complete evidence | ≥95% for production release |
| Change impact coverage | ECR items with linked affected objects | 100% for Class 1 changes |
| BOM accuracy | M-BOM match to released E-BOM | Zero critical mismatches |
| Cycle time | Req → approved release | Target −20% vs baseline |

---

## Slide 9 — Risks & mitigations

| Risk | Likelihood | Mitigation |
|------|------------|------------|
| Legacy data migration | High | Phased object scope; cleansing sprint |
| Authoring vs PLM split | Medium | Hybrid architecture clarity; agent bridge |
| User adoption | Medium | Champion network; template libraries |
| Integration overload | Medium | MVP boundary; API-first interfaces |

---

## Slide 10 — Recommended next step

**10-page concept alignment workshop + MVP charter (6 weeks)**

Deliverables:
- Validated process map for one pilot block/hull
- Object model and integration sketch
- Executive one-pager for funding decision

**Contact:** leejunha781@gmail.com | LinkedIn: joonha-lee-20b518316

---

## Build instructions (internal)

To produce full PPTX from this outline:

```powershell
# Use existing PLM builders — copy canonical deck first
$src = "E:\이력서\AVEVA - Marine Principal Technical Support & Consultant – PLM SME, Busan\Proposal\Future_Industrial_PLM_Meeting_Deck_EN.pptx"
# Or build fresh 10-slide deck via plm_slide_work\build_summary_en.ps1 pattern
```

Reference memory: `.cursor/memory/aveva-plm-application-deck.md`
