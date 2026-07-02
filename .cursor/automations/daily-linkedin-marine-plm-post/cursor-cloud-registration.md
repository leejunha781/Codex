# Register Daily LinkedIn Marine PLM Post in Cursor Automations

Use this guide to register the automation in **Cursor Desktop / cursor.com/automations** alongside the existing Codex automation.

## 1. Sync repo files to local Cursor automations dir (Windows)

```powershell
cd <repo-root>\.cursor\automations
.\sync-daily-linkedin-automation.ps1
.\validate-daily-linkedin-automation.ps1
```

Target: `C:\Users\namma\.cursor\automations\daily-linkedin-marine-plm-post\`

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
| Post + image output | `C:\Users\namma\Documents\Codex\YYYY-MM-DD\<topic-slug>\` | Repo `.cursor/automations/.../runs/YYYY-MM-DD/<topic-slug>/` |
| LinkedIn composer | Windows LinkedIn app (full-access sandbox) | Computer use if Windows env available; else report **ready for final posting** |
| Final Post click | User confirmation required | User confirmation required |

## 4. Activation checklist

- [ ] Scheduled trigger: daily 09:00
- [ ] Repository connected: `leejunha781/Codex`
- [ ] Memories enabled
- [ ] Computer use enabled
- [ ] Prompt matches [prompt.md](prompt.md)
- [ ] First run reviewed in Cursor Automations inbox
- [ ] Codex automation remains ACTIVE for local LinkedIn posting when PC is on

## 5. Keep Codex and Cursor in sync

After each run, topic rotation should be visible in:

- `.cursor/automations/daily-linkedin-marine-plm-post/memory.md`
- `.codex/automations/daily-linkedin-marine-plm-post/memory.md` (Codex local)
- Cursor Memories for this automation (UI)

Run both sync scripts after pulling repo updates:

```powershell
.\.codex\automations\sync-daily-linkedin-automation.ps1
.\.cursor\automations\sync-daily-linkedin-automation.ps1
```
