# FAT / SAT Checklist — LEO Maritime VSAT Terminal

**Document ID:** FAT-SAT-MTX4500-v1.0  
**Equipment:** LEO Maritime VSAT Terminal (Model MTX-4500)  
**Prepared by:** Joonha Lee, Engineering Documentation Services  
**Date:** 2026-07-09  
**Type:** Sample portfolio document (anonymized)

---

## 1. Document Control

| Version | Date | Author | Change |
|---------|------|--------|--------|
| 1.0 | 2026-07-09 | J. Lee | Initial sample for portfolio |

---

## 2. Scope

This checklist covers Factory Acceptance Test (FAT) at supplier facility and Site Acceptance Test (SAT) on vessel or designated site. It is intended for commissioning engineers, QA, and customer witnesses.

**In scope:** Antenna assembly, modem, power, network, pointing, link stability, documentation handover.  
**Out of scope:** Satellite operator NOC provisioning (separate operator checklist).

---

## 3. Roles

| Role | Responsibility |
|------|----------------|
| Supplier lead | Execute tests, record results |
| Customer witness | Observe critical tests, sign acceptance |
| Vessel electrician | Power distribution support (SAT) |

---

## 4. FAT Checklist — Factory

| ID | Test item | Method / criterion | Result | Pass/Fail | Notes |
|----|-----------|-------------------|--------|-----------|-------|
| F-01 | Visual inspection | No mechanical damage; labels correct | | ☐ Pass ☐ Fail | |
| F-02 | Power-on self-test | All LEDs per startup sequence | | ☐ Pass ☐ Fail | |
| F-03 | Input voltage range | 22–30 VDC bench supply | | ☐ Pass ☐ Fail | Record min/max |
| F-04 | Modem firmware version | Match release note | | ☐ Pass ☐ Fail | Ver: ______ |
| F-05 | Antenna pointing simulation | Pointing algo completes <3 min | | ☐ Pass ☐ Fail | Lab simulator |
| F-06 | Link establishment | Lock + data session 30 min | | ☐ Pass ☐ Fail | |
| F-07 | Throughput test | ≥ contracted Mbps down/up | | ☐ Pass ☐ Fail | |
| F-08 | Alarm reporting | Inject fault; NMS receives alarm | | ☐ Pass ☐ Fail | |
| F-09 | Documentation package | Manual, drawings, certs complete | | ☐ Pass ☐ Fail | |
| F-10 | Spare parts kit | Per BOM | | ☐ Pass ☐ Fail | |

**FAT overall:** ☐ Accept ☐ Reject — Signature: _______________ Date: _______

---

## 5. SAT Checklist — Site / Vessel

| ID | Test item | Method / criterion | Result | Pass/Fail | Notes |
|----|-----------|-------------------|--------|-----------|-------|
| S-01 | Mechanical installation | Mount, cable routing per drawing | | ☐ Pass ☐ Fail | |
| S-02 | Ground resistance | Radome base <1 Ω | | ☐ Pass ☐ Fail | Measured: ___ Ω |
| S-03 | DC supply at terminal | 22–30 V under steady load | | ☐ Pass ☐ Fail | |
| S-04 | **Power dip under load step** | Min voltage ≥22 V during engine/load step | | ☐ Pass ☐ Fail | *Added after RCA lesson* |
| S-05 | Initial satellite acquisition | Lock within spec time | | ☐ Pass ☐ Fail | |
| S-06 | 4 h link stability | No unplanned dropouts | | ☐ Pass ☐ Fail | |
| S-07 | 48 h link stability (optional) | Customer witness if required | | ☐ Pass ☐ Fail | |
| S-08 | Crew training | Basic ops + alarm response | | ☐ Pass ☐ Fail | |
| S-09 | As-built configuration record | Serials, IP, firmware logged | | ☐ Pass ☐ Fail | |
| S-10 | Evidence pack delivered | Logs, photos, signed checklist | | ☐ Pass ☐ Fail | |

**SAT overall:** ☐ Accept ☐ Reject — Signature: _______________ Date: _______

---

## 6. Defect Log

| Defect ID | FAT/SAT | Description | Severity | Disposition | Closed |
|-----------|---------|-------------|----------|-------------|--------|
| | | | | | |

---

## 7. Acceptance Sign-off

| Party | Name | Signature | Date |
|-------|------|-----------|------|
| Supplier | | | |
| Customer | | | |

---

*Portfolio sample. For client delivery, export to Word via Office COM or use `freelance/templates/FAT_SAT_Checklist_Template.md`.*
