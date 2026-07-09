# LinkedIn Reference Style — Commissioning Gates Solution Overview

**CANONICAL REFERENCE.** All future LinkedIn infographic images (Cursor Pro + Codex Pro) must match this professional graphic design standard.

**Verified reference run:** `leo-satellite-terminal-commissioning-gates` (2026-07-06)  
**Style name:** Commissioning Gates Solution Overview

---

## What this looks like

A high-end B2B technical whitepaper page — dark-mode corporate aesthetic (Linear/Vercel/Figma design language). Information-dense but grid-disciplined. **NOT** a simple AI poster, clipart, or decorative infographic.

---

## Mandatory layout (all sections required)

### Header
- **Main title:** ALL CAPS, bold white, large (topic-specific, e.g. "LEO SATELLITE TERMINAL: COMMISSIONING GATES")
- **Subtitle:** cyan/light blue, insight line (e.g. "Evidence before sea acceptance — not install-and-hope")
- **Badge:** "SOLUTION OVERVIEW" chip, top-right, blue border

### Top row — 7 numbered gate cards
Each card MUST contain:
- Number + icon (minimalist white line-art)
- **Bold gate title** (e.g. "COMMISSIONING REQUEST", "MODEM LOCK VALIDATION")
- **2–3 bullet points** in lighter cyan with specific technical detail (scope, thresholds, records)

Connect cards with small arrows. Equal width, aligned baselines.

### Center — photorealistic scene + technical overlay
- **Main visual:** high-quality photorealistic maritime/defense image (naval vessel at port/sea, twilight lighting) — NOT flat cartoon
- **HUD overlay pipeline:** horizontal flow connecting labeled nodes (e.g. Site Survey → Antenna → Modem → Shore Gateway → Ops Center)
- **TRACED LINK:** solid cyan line from overlay icon to actual hardware on the photo (e.g. antenna on ship)
- **HOLD / ESCALATE:** amber dashed line + warning triangle only for exception/hold path — restrained orange use

### Left sidebar — event/status legend
- Section title (e.g. "COMMISSIONING EVENTS")
- Vertical checklist with status icons (survey complete, link locked, failover tested, acceptance signed)

### Right sidebar — automation panel
- Section title (e.g. "PYTHON AUTOMATION")
- Code icon `</>` header
- 4 automation tasks with small icons (parse logs, compare thresholds, flag holds, export evidence)

### Bottom row — 3 value pillars
Numbered cards with icon + title + one-line description:
1. Traceable Baseline (shield icon)
2. Auditable Acceptance (scales/audit icon)
3. Export Evidence (document icon)

### Footer strip
- Target/bullseye icon + one-sentence engineering takeaway
- Example: "Automation widens test coverage — accountable commissioning keeps the operational baseline defensible."

---

## Color and typography

| Element | Spec |
|---------|------|
| Background | Deep navy `#0A1628` – `#0D1B2A` |
| Title | White, bold geometric sans (Inter-like) |
| Subtitle / bullets | Cyan `#00B4D8` – `#4CC9F0` |
| Body | White or light gray |
| Warning/hold only | Amber/orange `#F59E0B` — nowhere else |
| Icons | White thin-line outline, consistent stroke |

---

## Quality bar vs BLOCKED

| PASS (post) | BLOCKED (re-render) |
|-------------|---------------------|
| All 7 sections present | Missing pipeline, sidebars, or footer |
| Gate cards have bullets + detail | Icon-only cards with no technical bullets |
| Photorealistic center scene | Flat cartoon-only center |
| TRACED LINK on real hardware | Floating lines with no clear target |
| Information-dense whitepaper feel | Generic AI template / decorative fluff |
| Executive/customer-review ready | Childish, blurry, or sloppy typography |

---

## Tool workflow (Figma → Notion → Linear → Claude)

| Tool | Role |
|------|------|
| **Figma** | Primary design — populate template, export PNG 1080×1350+ |
| **Notion** | Calendar topic + Design Grade (Pass/Fail/Rework) + Figma URL |
| **Linear** | Issue `LinkedIn Design QA — <topic-slug>` in LinkedIn Content Pipeline |
| **Claude** | 09:20 QA — compare output to THIS reference; block if not match |
| **Cursor Pro** | 09:00 cloud — draft post + professional image |
| **Codex Pro** | Local backup generation + mirror/post 09:35 |

---

## design-brief.md template (per run)

```markdown
# Design Brief — <topic-slug>
Reference style: commissioning-gates-solution-overview

## Header
- Title (ALL CAPS):
- Subtitle (cyan insight line):

## 7 Gates (number, title, 2-3 bullets each)
1.
2.
...
7.

## Center scene
- Photorealistic subject:
- Overlay nodes:
- TRACED LINK target on photo:
- HOLD/ESCALATE exception path:

## Left sidebar (4 status items)
## Right sidebar (4 automation tasks)
## Bottom pillars (3)
## Footer takeaway
```

---

## Image Gen prompt anchor (when Figma unavailable)

When using Image Gen, include ALL of:
- "Match LEO SATELLITE TERMINAL COMMISSIONING GATES solution-overview reference style"
- "Photorealistic naval vessel center with technical HUD overlay"
- "7 gate cards with titles AND bullet points"
- "TRACED LINK cyan line to ship antenna"
- "HOLD/ESCALATE amber dashed warning"
- "Left commissioning events sidebar, right Python automation sidebar"
- "3 bottom value pillars, footer takeaway with target icon"
- "Deep navy dark mode, NOT cartoon, NOT generic AI infographic"
