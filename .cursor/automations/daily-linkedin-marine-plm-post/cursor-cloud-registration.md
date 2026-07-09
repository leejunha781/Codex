# Register Daily LinkedIn Marine PLM Post in Cursor Automations

Use this guide to register the automation in **Cursor Desktop / cursor.com/automations** alongside the existing Codex automation.

## 1. Sync repo files to local Cursor automations dir (Windows)

If `sync-both-linkedin-automations.ps1` or `validate-daily-linkedin-automation.ps1` are missing under `C:\Users\namma\.cursor\automations\`, files are not on your current git branch (main/memory). Use GitHub fetch instead:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -Command "iwr -UseBasicParsing 'https://raw.githubusercontent.com/leejunha781/Codex/cursor/fix-linkedin-local-mirror-0681/.cursor/automations/fetch-cursor-linkedin-scripts-from-github.ps1' -OutFile $env:TEMP\fetch-linkedin.ps1; powershell -NoProfile -ExecutionPolicy Bypass -File $env:TEMP\fetch-linkedin.ps1"
```

Or after switching branch:

```powershell
cd C:\Users\namma
git fetch origin cursor/fix-linkedin-local-mirror-0681
git checkout cursor/fix-linkedin-local-mirror-0681
git pull origin cursor/fix-linkedin-local-mirror-0681
powershell -NoProfile -File .cursor\automations\install-cursor-linkedin-automation-scripts.ps1
```

Then:

```powershell
cd C:\Users\namma\.cursor\automations
.\sync-both-linkedin-automations.ps1
.\validate-daily-linkedin-automation.ps1
```

Required files in `C:\Users\namma\.cursor\automations\`:
- `sync-both-linkedin-automations.ps1`
- `sync-daily-linkedin-automation.ps1`
- `validate-daily-linkedin-automation.ps1`
- `mirror-linkedin-runs.ps1`
- `post-linkedin-windows-app-prompt.md`
- `install-cursor-linkedin-automation-scripts.ps1`

Target: `C:\Users\namma\.cursor\automations\daily-linkedin-marine-plm-post\`

Scripts synced to `C:\Users\namma\.cursor\automations\`:
- `mirror-linkedin-runs.ps1`
- `post-linkedin-windows-app-prompt.md`

## 2. Create automation in Cursor UI

1. Open [cursor.com/automations/new](https://cursor.com/automations/new) or **Agents Window → Automations → New**
2. **Name:** `Daily LinkedIn Marine PLM Post`
3. **Trigger:** Scheduled — daily at **09:00** (cron: `0 9 * * *`)
4. **Repository:** `leejunha781/Codex` (branch: `master`)
5. **Tools:** enable **Memories** and **Computer use**
6. **Permissions:** Private (or Team Visible if shared)
7. **Prompt:** paste contents of [prompt.md](prompt.md) in this folder

## 3. Cursor vs Codex execution split

| Step | Codex (local) | Cursor (cloud) |
|------|-----------------|----------------|
| Topic rotation | Notion Content Calendar + `memory.md` | Notion MCP query + `memory.md` + Cursor Memories |
| Post + image output | `C:\Users\namma\Documents\Codex\YYYY-MM-DD\<topic-slug>\` (direct, backup automation) | Repo `.cursor/automations/.../runs/YYYY-MM-DD/<topic-slug>/` |
| Infographic design | Image Gen fallback (local) | Figma MCP (preferred) or Image Gen |
| QA review | `daily-linkedin-claude-review` at 09:20 | Notion status → QA Review |
| Windows mirror | `C:\Users\namma\Documents\Codex\YYYY-MM-DD\<topic-slug>\` via `mirror-linkedin-runs.ps1` | Not available in cloud |
| LinkedIn posting | `daily-linkedin-mirror-and-post` at 09:35 — auto-post via Windows LinkedIn app | Not available in cloud |
| Archive | Notion status → Posted | Notion status → QA Review |

## 3a. Figma + Notion + Claude integration

See [figma-notion-claude-integration.md](figma-notion-claude-integration.md) for full guide.

- **Notion Content Calendar:** https://app.notion.com/p/7ee8f488584e4ca290f3fcbfa3ea1314
- **Config:** `notion-config.json`
- **Claude QA:** `daily-linkedin-claude-review` at 09:20

Enable **Notion** and **Figma** MCP in Cursor automation tools (alongside Memories and Computer use).

## 4. Local folder mirroring + auto-posting

**Fixed Windows mirror path:** `C:\Users\namma\Documents\Codex\YYYY-MM-DD\<topic-slug>\`

Daily pipeline:

1. **09:00** — Cursor cloud: query Notion topic → generate artifacts in repo `runs/` → update Notion (QA Review)
2. **09:20** — Codex `daily-linkedin-claude-review`: post + image QA → Notion status Ready or Blocked
3. **09:35** — Codex `daily-linkedin-mirror-and-post`:
   - Mirrors to `C:\Users\namma\Documents\Codex\YYYY-MM-DD\<topic-slug>\`
   - Auto-posts via LinkedIn Windows app (no user confirmation)
   - Updates Notion status to Posted
3. Files in mirror folder:
   - `linkedin-post.md`
   - `<topic-slug>-infographic.png`
   - `ready-for-posting.json`

Manual mirror only:

```powershell
cd <repo-root>\.cursor\automations
.\mirror-linkedin-runs.ps1 -Pull -IncludeRemoteBranches
```

## 5. Activation checklist

- [ ] Cursor cloud trigger: daily 09:00 (tools: Memories, Computer use, **Notion**, **Figma**)
- [ ] Notion MCP authenticated in Cursor Desktop
- [ ] Figma MCP authenticated in Cursor Desktop
- [ ] Notion Content Calendar seeded: https://app.notion.com/p/7ee8f488584e4ca290f3fcbfa3ea1314
- [ ] Codex `daily-linkedin-claude-review`: daily 09:20, ACTIVE
- [ ] Codex `daily-linkedin-mirror-and-post`: daily 09:35, ACTIVE, local execution, full-access sandbox
- [ ] Disable legacy `daily-linkedin-mirror-runs` if still present in Codex UI
- [ ] LinkedIn Windows app installed and logged in
- [ ] PC awake at 09:35 (enable "Prevent sleep while running")
- [ ] Mirror log: `.cursor\automations\cache\linkedin-mirror\mirror.log`

## 6. Keep Codex and Cursor in sync

**If `running scripts is disabled` error appears**, use `-ExecutionPolicy Bypass`:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File "$env:USERPROFILE\.cursor\automations\run-linkedin-automation-setup.ps1"
```

Or run individually:

```powershell
cd $env:USERPROFILE\.cursor\automations
powershell -NoProfile -ExecutionPolicy Bypass -File .\sync-both-linkedin-automations.ps1
powershell -NoProfile -ExecutionPolicy Bypass -File .\validate-daily-linkedin-automation.ps1
```

One-time fix (optional, CurrentUser only):

```powershell
Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned
```

This syncs automation definitions, mirror/post scripts, and runs an immediate mirror pass (no auto-post).
