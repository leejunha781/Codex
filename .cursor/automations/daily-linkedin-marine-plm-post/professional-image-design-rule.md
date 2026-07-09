# Professional Image Design Rule — LinkedIn Infographics

**MANDATORY.** Every LinkedIn post image must be a professionally designed engineering deliverable — never a simple AI poster, clipart slide, or generic infographic.

Read with: `figma-notion-claude-integration.md`, `notion-config.json`, `linear-config.json`.

---

## Non-negotiable quality bar

The final PNG must pass ALL of these:

| Criterion | Requirement |
|-----------|-------------|
| Layout | Solution-overview poster: 5–7 step pipeline + central scene + sidebars + 3 summary cards + footer |
| Typography | Clear hierarchy (title > section labels > body); no clipped or overflowing text |
| Icons | Flat monochrome-blue vector, consistent line-weight — not cartoon, not 3D emoji |
| Grid | Disciplined alignment, balanced whitespace, executive-grade composition |
| Color | Deep-navy `#0A1628` background; cyan labels; white body; amber only for alert/fallback |
| Resolution | Minimum 1080×1350 (LinkedIn portrait); export at 2× if possible |
| Feel | Senior information designer / consulting deliverable — NOT stock AI infographic |

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

- Title / subtitle (English)
- Pipeline steps (5–7): number, label, one-line description
- Central scene: what system/architecture to show
- Left legend items (3–5)
- Right sidebar: automation/tooling bullets (3–4)
- Bottom summary cards (3): title + one line each
- Footer takeaway strip (one sentence)
- Color notes: primary path cyan, exception path amber
```

Use this brief to drive Figma `use_figma` or Image Gen prompt.

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
