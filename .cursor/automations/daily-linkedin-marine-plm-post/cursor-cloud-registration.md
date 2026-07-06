# Register Daily LinkedIn Marine PLM Post in Cursor Automations

Use this guide to register the automation in **Cursor Desktop / cursor.com/automations** alongside the existing Codex automation.

## 1. Sync repo files to local Cursor automations dir (Windows)

```powershell
cd <repo-root>\.cursor\automations
.\sync-both-linkedin-automations.ps1
.\validate-daily-linkedin-automation.ps1
```

Target: `C:\Users\namma\.cursor\automations\daily-linkedin-marine-plm-post\`

Mirror script: `C:\Users\namma\.cursor\automations\mirror-linkedin-runs.ps1`

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
| Post + image output | `C:\Users\namma\Documents\Codex\YYYY-MM-DD\<topic-slug>\` (direct) | Repo `.cursor/automations/.../runs/YYYY-MM-DD/<topic-slug>/` |
| Windows mirror | `mirror-linkedin-runs.ps1` at 09:30 via `daily-linkedin-mirror-runs` | Not available in cloud; repo artifacts only |
| LinkedIn composer | Windows LinkedIn app (full-access sandbox) | Computer use if Windows env available; else report **ready for final posting** |
| Final Post click | User confirmation required | User confirmation required |

## 4. Local folder mirroring (cloud → Windows)

Cloud runs cannot write to `C:\Users\namma\Documents\Codex\`. Use the mirror pipeline:

1. **09:00** — Cursor cloud generates artifacts in repo `runs/`
2. **09:30** — Codex local automation `daily-linkedin-mirror-runs` runs:
   ```powershell
   powershell -NoProfile -File C:\Users\namma\.cursor\automations\mirror-linkedin-runs.ps1 -Pull -IncludeRemoteBranches
   ```
3. Output appears at `C:\Users\namma\Documents\Codex\YYYY-MM-DD\<topic-slug>\`

Manual mirror anytime:

```powershell
cd <repo-root>\.cursor\automations
.\mirror-linkedin-runs.ps1 -Pull -IncludeRemoteBranches
.\mirror-linkedin-runs.ps1 -AllDates -Pull -IncludeRemoteBranches   # backfill all dates
```

## 5. Activation checklist

- [ ] Scheduled trigger: daily 09:00 (Cursor cloud)
- [ ] Mirror automation: daily 09:30 (Codex local `daily-linkedin-mirror-runs`)
- [ ] Repository connected: `leejunha781/Codex`
- [ ] Memories enabled
- [ ] Computer use enabled
- [ ] Prompt matches [prompt.md](prompt.md)
- [ ] First run reviewed in Cursor Automations inbox
- [ ] Mirror log checked: `.cursor\automations\cache\linkedin-mirror\mirror.log`
- [ ] Codex automation remains ACTIVE for local LinkedIn posting when PC is on

## 6. Keep Codex and Cursor in sync

After each run, topic rotation should be visible in:

- `.cursor/automations/daily-linkedin-marine-plm-post/memory.md`
- `.codex/automations/daily-linkedin-marine-plm-post/memory.md` (Codex local)
- Cursor Memories for this automation (UI)

Run sync after pulling repo updates:

```powershell
.\.cursor\automations\sync-both-linkedin-automations.ps1
```

This syncs automation definitions, mirror script, and runs an immediate mirror pass.
