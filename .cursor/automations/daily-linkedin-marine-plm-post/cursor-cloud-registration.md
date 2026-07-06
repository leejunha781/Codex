# Register Daily LinkedIn Marine PLM Post in Cursor Automations

Use this guide to register the automation in **Cursor Desktop / cursor.com/automations** alongside the existing Codex automation.

## 1. Sync repo files to local Cursor automations dir (Windows)

```powershell
cd <repo-root>\.cursor\automations
.\sync-both-linkedin-automations.ps1
.\validate-daily-linkedin-automation.ps1
```

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
| Topic rotation | `~/.codex/automations/.../memory.md` | Repo `memory.md` + Cursor Memories |
| Post + image output | `C:\Users\namma\Documents\Codex\YYYY-MM-DD\<topic-slug>\` (direct, backup automation) | Repo `.cursor/automations/.../runs/YYYY-MM-DD/<topic-slug>/` |
| Windows mirror | `C:\Users\namma\Documents\Codex\YYYY-MM-DD\<topic-slug>\` via `mirror-linkedin-runs.ps1` | Not available in cloud |
| LinkedIn posting | `daily-linkedin-mirror-and-post` at 09:35 — auto-post via Windows LinkedIn app | Not available in cloud |

## 4. Local folder mirroring + auto-posting

**Fixed Windows mirror path:** `C:\Users\namma\Documents\Codex\YYYY-MM-DD\<topic-slug>\`

Daily pipeline:

1. **09:00** — Cursor cloud generates artifacts in repo `runs/`
2. **09:35** — Codex `daily-linkedin-mirror-and-post`:
   - Mirrors to `C:\Users\namma\Documents\Codex\YYYY-MM-DD\<topic-slug>\`
   - Auto-posts via LinkedIn Windows app (no user confirmation)
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

- [ ] Cursor cloud trigger: daily 09:00
- [ ] Codex `daily-linkedin-mirror-and-post`: daily 09:35, ACTIVE, local execution, full-access sandbox
- [ ] Disable legacy `daily-linkedin-mirror-runs` if still present in Codex UI
- [ ] LinkedIn Windows app installed and logged in
- [ ] PC awake at 09:35 (enable "Prevent sleep while running")
- [ ] Mirror log: `.cursor\automations\cache\linkedin-mirror\mirror.log`

## 6. Keep Codex and Cursor in sync

```powershell
.\.cursor\automations\sync-both-linkedin-automations.ps1
```

This syncs automation definitions, mirror/post scripts, and runs an immediate mirror pass (no auto-post).
