# Professional Image Design Rule — LinkedIn Infographics

**MANDATORY.** Every LinkedIn post image must be a professionally designed engineering deliverable — never a simple AI poster, clipart slide, or generic infographic.

Read with: `figma-notion-claude-integration.md`, `linkedin-reference-style-commissioning-gates.md`, `notion-config.json`, `linear-config.json`.

**Canonical visual reference:** `linkedin-reference-style-commissioning-gates.md` — LEO Satellite Terminal Commissioning Gates solution-overview (photorealistic vessel + 7 gate cards with bullets + HUD overlay + sidebars + value pillars). Cursor Pro and Codex Pro MUST match this standard.

---

## Non-negotiable quality bar

The final PNG must pass ALL of these:

| Criterion | Requirement |
|-----------|-------------|
| Layout | Commissioning Gates solution-overview: header + badge, 7 gate cards **with bullet details**, photorealistic center scene + HUD overlay, left event sidebar, right automation sidebar, 3 value pillars, footer |
| Center visual | Photorealistic maritime/defense photo (vessel, shipyard, ops room) with technical overlay — NOT flat cartoon only |
| Gate cards | Each of 7 cards: number, icon, bold title, 2–3 cyan bullet points with technical specifics |
| Overlay links | TRACED LINK (solid cyan to hardware on photo); HOLD/ESCALATE (amber dashed, warning icon only) |
| Typography | Clear hierarchy (title > section labels > body); no clipped or overflowing text |
| Icons | Flat monochrome-blue vector, consistent line-weight — not cartoon, not 3D emoji |
| Grid | Disciplined alignment, balanced whitespace, executive-grade composition |
| Color | Deep-navy `#0A1628` background; cyan labels; white body; amber only for alert/fallback |
| Resolution | Minimum 1080×1350 (LinkedIn portrait); export at 2× if possible |
| Feel | High-end B2B technical whitepaper / commissioning gates poster — NOT stock AI infographic |

## Reference style check (compare to commissioning-gates)

Before posting, verify image matches `linkedin-reference-style-commissioning-gates.md`:
- [ ] SOLUTION OVERVIEW badge present
- [ ] 7 gates with bullet-point detail (not icon-only)
- [ ] Photorealistic center scene with HUD overlay
- [ ] TRACED LINK line terminates on visible hardware
- [ ] Left status sidebar + right automation sidebar
- [ ] 3 bottom value pillars + footer takeaway

## BLOCKED — do not post if image shows any of these

- Childish, cartoon, or clipart style
- Generic "AI infographic" template look (random icons, no grid, decorative fluff)
- Plain text-on-color blocks with no architecture visual
- Single hero image with no structured information hierarchy
- Blurry, low-resolution, or obviously auto-generated sloppy typography
- Missing title, pipeline, or takeaway strip

If blocked → re-render. Maximum 3 attempts, then set Notion/Linear status **Blocked** with reason.

---

## Production pipeline (strict priority)

```
1. Figma MCP (PRIMARY)
   ├── search_design_system → reuse components/tokens
   ├── get_design_context → read template layout
   ├── use_figma → populate variables (title, pipeline, legend, sidebar, cards)
   └── download_assets → export PNG to runs/ folder

2. generate_figma_design (SECONDARY)
   └── Push structured layout into existing Figma file, then export

3. Built-in Image Gen (TERTIARY — only if Figma unavailable)
   └── MUST follow full solution-overview spec in prompt.md
   └── MUST re-render if result looks simple/generic (compare to reference-grade rule)

4. Pillow/PowerShell (LAST RESORT — requires explicit blocker note)
```

Record `Image Source` in Notion: `Figma Template` | `Figma Generated` | `Image Gen` | `Pillow Fallback`.

---

## Per-run design brief (create before image)

Save `design-brief.md` in each run folder:

```markdown
# Design Brief — <topic-slug>

- Title / subtitle (English) + SOLUTION OVERVIEW badge
- Pipeline gates (7): number, bold title, 2-3 bullet points each (technical detail)
- Central scene: photorealistic subject + overlay nodes + TRACED LINK + HOLD/ESCALATE path
- Left legend: 4 commissioning/status events with checkmarks
- Right sidebar: Python automation tasks (4) with code icon header
- Bottom value pillars (3): icon + title + one line each
- Footer takeaway strip with target icon (one sentence)
- Reference style: commissioning-gates-solution-overview
```

Use this brief to drive Figma `use_figma` or Image Gen prompt.

---

## C2PA / Content Credentials strip (MANDATORY before LinkedIn upload)

LinkedIn displays a **"Content credentials" (CR)** badge when PNG/JPEG files contain **C2PA metadata** (often embedded by AI image tools such as DALL·E / Image Gen as a `caBX` PNG chunk).

| Step | Action |
|------|--------|
| After image export | Strip C2PA metadata before saving final PNG |
| Windows mirror | `mirror-linkedin-runs.ps1` auto-strips `*-infographic.png` after copy |
| Manual strip | `.\strip-linkedin-image-c2pa.ps1 -ImagePath "...\topic-infographic.png"` |
| Cloud/Linux | `python3 strip-linkedin-image-c2pa.py path/to/*-infographic.png` |

**QA gate:** FAIL if `caBX`, `c2pa`, or `jumb` strings remain in the PNG bytes. Re-export clean PNG before posting.

**Figma export:** If using Adobe/Figma Content Credentials, disable "Apply Content Credentials" on export when possible.

---

## Tool integration per run

### Notion
- Set `Image Source` and `Image QA` on calendar row
- Add `Design Grade`: Pass | Fail | Rework
- Link `Figma File` URL when used

### Figma
- Always attempt Figma MCP before Image Gen
- If View-seat blocks edit: use `generate_figma_design` or document blocker in Linear
- Target: reusable LinkedIn solution-overview template file

### Linear
- Create/update issue per run: `LinkedIn Design QA — <topic-slug>`
- Project: `LinkedIn Content Pipeline`
- States: Design Brief → In Design → QA Review → Approved | Blocked
- Attach Figma URL + repo run path in issue description
- Claude QA at 09:20 updates issue to Approved or Blocked

---

## Reference layout (solution-overview)

```
┌─────────────────────────────────────────────────────────┐
│  TITLE + SUBTITLE                                       │
├─────────────────────────────────────────────────────────┤
│  [1] [2] [3] [4] [5] [6] [7]  ← pipeline icon cards    │
├──────────┬──────────────────────────────┬───────────────┤
│  LEGEND  │   CENTRAL ARCHITECTURE       │  AUTOMATION   │
│  panel   │   SCENE (vessel/sat/PLM)     │  SIDEBAR      │
├──────────┴──────────────────────────────┴───────────────┤
│  [Summary 1]    [Summary 2]    [Summary 3]              │
├─────────────────────────────────────────────────────────┤
│  FOOTER TAKEAWAY STRIP                                  │
└─────────────────────────────────────────────────────────┘
```

This matches the Marine PLM solution-overview reference style from prior verified runs (satcom-fallback, BOM validation gates).
