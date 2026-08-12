# Figma + Notion + Linear + Claude LinkedIn Integration

Operational guide for the Daily LinkedIn Marine PLM Post pipeline using **Notion** (content calendar), **Figma** (professional infographic design), **Linear** (design QA tracking), and **Claude** (QA review).

**Image quality:** Read `professional-image-design-rule.md` — all images must be professionally designed, never simple AI posters.

## Pipeline overview

```mermaid
flowchart LR
    N["Notion Calendar\n(topic + angle)"] --> D["Design Brief\n+ Figma MCP"]
    D --> C["Cursor Cloud 09:00\n(draft + pro image)"]
    C --> L["Linear\nDesign QA issue"]
    L --> CL["Claude Review 09:20\n(QA gate)"]
    CL --> M["Codex Mirror 09:35\n+ LinkedIn post"]
    M --> N2["Notion archive\n(status=Posted)"]
```

| Time | Automation | Tool | Action |
|------|-----------|------|--------|
| 09:00 | Cursor Cloud `daily-linkedin-marine-plm-post` | Notion, **Figma MCP**, Linear | Query topic → design-brief → **Figma professional image** → save repo → Notion QA Review → Linear issue |
| 09:20 | Codex `daily-linkedin-claude-review` | Claude (local) | Post + **design grade** QA → Notion/Linear Approved or Blocked |
| 09:35 | Codex `daily-linkedin-mirror-and-post` | PowerShell + LinkedIn app | Mirror → post → Notion Posted |

---

## Notion — Content Calendar

**Database:** [LinkedIn Content Calendar](https://app.notion.com/p/7ee8f488584e4ca290f3fcbfa3ea1314)

**Config file:** `notion-config.json` (same folder)

### Properties

| Property | Purpose |
|----------|---------|
| Topic | Display title |
| Topic Slug | Folder name (`YYYY-MM-DD/<topic-slug>/`) |
| Scheduled Date | Calendar date for posting |
| Angle | One-line angle for the post |
| Marine Context | Execution context (BOM, SATCOM, shipyard gates, etc.) |
| Status | Planned → Drafting → QA Review → Ready → Posted / Blocked |
| Image Source | Figma Template / Image Gen / Pillow Fallback |
| Image QA | QA checklist results |
| Repo Path | GitHub run folder URL |
| Windows Path | `C:\Users\namma\Documents\Codex\...` |
| Figma File | Figma template or export URL |
| LinkedIn Status | posted / ready / blocked |
| Notes | Free-form notes |

### Cursor Cloud — pre-run (09:00)

1. Read `notion-config.json` for database IDs.
2. Query Notion for today's topic:
   - SQL: `query_today_sql` with today's date (`YYYY-MM-DD`)
   - Fallback: `query_next_planned_sql` if no topic scheduled today
3. If no Notion result, fall back to `memory.md` topic rotation.
4. Set Notion status to **Drafting** when run starts.

### Cursor Cloud — post-run (09:00)

1. Set Notion status to **QA Review** only. Do **not** set **Ready** — Claude QA (`daily-linkedin-claude-review` at 09:20) is the sole path that may promote status to Ready.
2. Fill: Image Source, Image QA, Repo Path, Windows Path, Figma File (if used).
3. Append run summary to `memory.md`.

### Codex mirror — post-run (09:35)

1. After LinkedIn post (or block), set Notion status to **Posted** or **Blocked**.
2. Fill LinkedIn Status field.

### Freelance topic rotation

Reference: `freelance-topic-reference.md`

- ~1 in 4 posts use a freelance service angle (Doc Automation, Resume/LinkedIn, PLM Planning, Automation Tools)
- Notion calendar includes freelance Planned topics (2026-07-13 through 2026-07-16)
- Posts remain consultative, not promotional

---

## Figma — Infographic Design

**Account:** 이준하 (leejunha781@gmail.com) — View seat on Starter plan.

### Design template layout (solution-overview reference)

Reference-grade flat vector infographic:

- **Background:** uniform deep-navy (`#0A1628`)
- **Top:** 5–7 numbered process pipeline icon cards
- **Center:** stylized system/architecture scene (vessel, satellite, PLM nodes)
- **Left:** legend panel (operational events, status codes)
- **Right:** automation/Python sidebar
- **Bottom:** 3 summary cards + footer takeaway strip
- **Typography:** cyan section labels, white body text, amber for alert/fallback paths only
- **Icons:** flat monochrome-blue vector, consistent thin line-weight

### Image production priority (Figma-first — MANDATORY)

See `professional-image-design-rule.md` for full spec.

1. **Figma MCP** (PRIMARY — always attempt first):
   - `search_design_system` → reuse components/tokens
   - `get_design_context` → read template
   - `use_figma` → populate title, pipeline, legend, sidebar, cards
   - `download_assets` → export PNG 1080×1350+
2. **generate_figma_design** (when template file exists)
3. **Built-in Image Gen** (ONLY if Figma blocked — must pass anti-simple-image check, up to 3 re-renders)
4. **Pillow/PowerShell** (last resort — document blocker)

**Never post** a simple/generic/childish image. Block and re-render instead.

### View-seat limitation

Current Figma seat is **View only**. To use Figma as primary image source:
- Upgrade to Edit seat, OR
- Share template file with edit access to the MCP-connected account
- Until then: use Image Gen with solution-overview style rules (proven in test runs)

### Figma file placeholder

When a template file is created, record its URL in:
- `notion-config.json` → add `figma_template_url`
- Notion calendar row → Figma File property

---

## Linear — Design QA Tracking

**Project:** [LinkedIn Content Pipeline](https://linear.app/joonha-lee/project/linkedin-content-pipeline-bf77c4a8e3b6)

**Config:** `linear-config.json`

### Per-run workflow

1. **Before image:** Create Linear issue `LinkedIn Design QA — <topic-slug>` in project LinkedIn Content Pipeline (team Joonha_Lee)
2. **After image:** Set issue state to QA Review; link Figma URL, Notion page, repo path
3. **Claude QA (09:20):** Update issue to Approved (pass) or Blocked (fail) with design grade notes
4. **Blocked images:** Do not proceed to mirror+post until re-designed

### Issue checklist (in description)

- Solution-overview layout present
- Professional typography and grid
- Not childish/cartoon/generic AI
- Text-fit, leader-lines, overlap QA pass

---

## Claude — QA Review Gate

**Automation:** `.codex/automations/daily-linkedin-claude-review/` (09:20 daily)

**Prompt:** `claude-review-prompt.md`

### Review checklist

Claude reads the latest run from repo `runs/YYYY-MM-DD/<topic-slug>/` and checks:

1. **Post quality:** consultative tone, marine execution context, 3 application points, 5–8 hashtags
2. **Image QA:** text-fit, leader-lines, overlap, photo-diversity (8/8 if applicable)
3. **Professional grade:** reference-grade flat vector style, not childish/cartoon/**not simple/generic AI poster**
4. **Design grade:** Pass only if image meets `professional-image-design-rule.md` (solution-overview layout, executive composition). Fail → Blocked, update Linear issue.
5. **Notion sync:** update status to Ready (pass) or Blocked (fail with reason)

### Memory sync

After review, append to:
- `.codex/automations/daily-linkedin-marine-plm-post/memory.md`
- Claude `MEMORY.md` (via `memory-sync-manual` if needed)

---

## Activation checklist

- [ ] Notion MCP authenticated in Cursor Desktop
- [ ] Figma MCP authenticated in Cursor Desktop
- [ ] Linear MCP authenticated in Cursor Desktop
- [ ] Claude Desktop running with git autosync on Windows
- [ ] Cursor automation tools: Memories, Computer use, **Notion**, **Figma**, **Linear**
- [ ] Codex `daily-linkedin-claude-review`: daily 09:20, ACTIVE
- [ ] Codex `daily-linkedin-mirror-and-post`: daily 09:35, ACTIVE
- [ ] LinkedIn Content Calendar seeded with upcoming topics

## Sync commands (Windows)

```powershell
cd $env:USERPROFILE\.cursor\automations
.\sync-both-linkedin-automations.ps1
.\validate-daily-linkedin-automation.ps1
```
