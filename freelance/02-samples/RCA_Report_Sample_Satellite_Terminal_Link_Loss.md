# Root Cause Analysis Report

**Document ID:** RCA-2026-SAT-0142  
**Classification:** Customer-facing technical report  
**Prepared by:** Joonha Lee, Engineering Documentation Services  
**Date:** 2026-07-09  
**Status:** Closed — corrective action verified

---

## 1. Executive Summary

A maritime LEO satellite communication terminal experienced intermittent link loss during sea trials. Field logs showed repeated modem re-registration events correlated with vessel motion and power rail dips. Root cause was identified as inadequate DC input filtering on the modem power feed combined with marginal RF cable grounding at the radome base. Corrective actions included installation of a marine-rated DC filter, rework of the ground strap, and a repeat SAT link-stability test. Post-correction testing achieved 48 hours continuous link with zero unplanned dropouts.

---

## 2. Issue Information

| Field | Detail |
|-------|--------|
| Equipment | LEO Maritime VSAT Terminal (Model MTX-4500) |
| Customer / Vessel | [Anonymized] — Offshore support vessel |
| Reported date | 2026-06-12 |
| Severity | High — mission communication impact |
| Symptom | Intermittent link loss; modem re-registration every 15–45 min |

---

## 3. Problem Description

The customer reported that the terminal established initial satellite lock successfully but lost connectivity repeatedly during vessel operation. Symptoms included:

- LED status cycling from **LINK OK** to **SEARCHING**
- NMS logs showing `MODEM_REREG` events without corresponding RF obstruction
- Crew observation that events increased during heavy seas and main-engine load changes

No physical damage to the radome or antenna was observed at first inspection.

---

## 4. Environment

| Parameter | Value |
|-----------|-------|
| Location | Open sea, East China Sea trial area |
| Vessel motion | Roll ±12°, pitch ±8° (moderate sea state) |
| Power source | Ship 24 VDC distribution panel, Branch CB-14 |
| Ambient temperature | 28–34 °C (enclosure internal 42 °C peak) |
| Installation age | 6 weeks post-commissioning |

---

## 5. Investigation & Test Evidence

### 5.1 Log review

| Timestamp (UTC) | Event | Notes |
|-----------------|-------|-------|
| 2026-06-12 03:14 | LINK_LOST | Preceded by `VIN_DIP 21.8V` |
| 2026-06-12 03:14:02 | MODEM_REREG | Recovery in 38 s |
| 2026-06-12 05:41 | LINK_LOST | `VIN_DIP 22.1V`, roll 11° |
| 2026-06-12 08:22 | LINK_LOST | No obstruction alarm |

### 5.2 On-site measurements

| Test | Result | Limit / Expectation |
|------|--------|---------------------|
| DC input at modem (steady state) | 24.2 V | 22–30 V |
| DC input during engine load step | 21.6 V (200 ms dip) | >22 V recommended |
| Radome base ground resistance | 4.8 Ω | <1 Ω per install guide |
| RF cable continuity | Pass | — |
| Blockage / obstruction sensor | No active alarms | — |

### 5.3 Failure symptom classification

- **Category:** Power integrity + grounding (installation-related)
- **Not primary:** Antenna pointing, modem firmware defect, or satellite network outage

---

## 6. Root Cause Analysis

### 6.1 Root cause (confirmed)

**Inadequate DC input filtering** on the modem power feed allowed transient undervoltage during ship load steps, triggering modem reset and link loss.

**Contributing factor:** **Marginal RF ground strap** at radome base (4.8 Ω) increased susceptibility to noise during motion, though primary trigger was power dip.

### 6.2 Why it was not caught at FAT

- Factory FAT used stable bench power supply; vessel distribution transients were not simulated.
- Ground strap torque was not verified with ohmmeter in SAT checklist (visual check only).

### 6.3 5-Why summary

1. Why link loss? → Modem re-registered repeatedly.  
2. Why re-register? → Power or RF instability.  
3. Why power instability? → 21.6–21.8 V dips on load step.  
4. Why dips? → No marine DC filter; long cable run from panel.  
5. Why missed? → SAT checklist did not include load-step power test.

---

## 7. Corrective Actions

| ID | Action | Owner | Target date | Status |
|----|--------|-------|-------------|--------|
| CA-01 | Install marine-rated DC filter (Model DF-24M) at modem input | Field service | 2026-06-18 | Complete |
| CA-02 | Replace and torque ground strap; verify <1 Ω | Field service | 2026-06-18 | Complete |
| CA-03 | Update SAT checklist: power dip test under load step | QA / docs | 2026-06-25 | Complete |
| CA-04 | Customer advisory bulletin for similar installations | Engineering | 2026-07-01 | Complete |

---

## 8. Verification

| Test | Procedure | Result | Pass/Fail |
|------|-----------|--------|-----------|
| V-01 | 48 h continuous link test at sea | 0 unplanned dropouts | **Pass** |
| V-02 | Power dip test with engine load step | Min 23.1 V at modem | **Pass** |
| V-03 | Ground resistance at radome base | 0.6 Ω | **Pass** |
| V-04 | NMS log review post-fix | No MODEM_REREG events | **Pass** |

**Verification sign-off:** Field test lead, 2026-06-22.

---

## 9. Preventive Recommendations

1. Add **load-step DC holdup test** to all maritime SAT procedures.  
2. Specify **maximum ground resistance** in installation work instructions.  
3. Include **power quality snapshot** in commissioning evidence pack.

---

## 10. Attachments

- A: NMS event log excerpt (2026-06-12)  
- B: Before/after ground strap photos  
- C: SAT re-test record (48 h link stability)

---

*This sample is anonymized for portfolio use. Structure aligns with Service 1 deliverable templates in `freelance/templates/`.*
