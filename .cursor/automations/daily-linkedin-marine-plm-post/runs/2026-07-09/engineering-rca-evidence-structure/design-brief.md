# Design Brief — engineering-rca-evidence-structure

**Date:** 2026-07-09  
**Image target:** 1080×1350 LinkedIn portrait PNG

## Title / Subtitle

- **Title:** RCA Evidence Chain — From Field Symptom to Customer Closure
- **Subtitle:** Engineering Document Automation for Defense & Satellite Programmes

## Pipeline (7 steps)

| # | Label | Description |
|---|-------|-------------|
| 1 | Intake | Field notes, photos, initial symptom log |
| 2 | Symptom Log | Time-stamped event sequence |
| 3 | Test Evidence | BER/SNR, terminal logs, measurements |
| 4 | Environment | Vessel, rack, network, mission context |
| 5 | Root Cause | Hypothesis + verification proof |
| 6 | Corrective Action | Fix + ownership assignment |
| 7 | Closure Pack | Customer-ready RCA export |

## Central scene

Satellite terminal on vessel deck linked to shore gateway. Cyan primary data path from terminal → gateway → evidence store. Amber dashed exception/fallback path for missing evidence links.

## Left legend — Evidence Types

- RF Metrics (BER/SNR)
- Gateway Logs
- Terminal State
- Mission Context
- Weather / Environment

## Right sidebar — Python Automation

- Normalize readings & timestamps
- Parse gateway / terminal logs
- Threshold & acceptance checks
- Export DOCX/PDF closure pack

## Bottom summary cards

1. **Traceability** — Every claim linked to captured artifact
2. **Review Gates** — Engineer sign-off before export
3. **PLM Handoff** — Structured pack for acceptance records

## Footer takeaway

AI drafts faster — engineers own the evidence chain.

## Color notes

- Background: `#0A1628` deep navy
- Labels: cyan `#00B4D8`
- Body text: white
- Alert/missing path: amber `#F59E0B` only
- Icons: flat monochrome blue, thin line-weight
