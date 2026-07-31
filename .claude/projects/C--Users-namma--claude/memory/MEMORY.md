# Memory Index

Canonical hub: `C:\Users\namma\.cursor\memory`. Synced across Claude, Codex, ChatGPT mirror, and Cursor.

## Entries
- [Airpoint FPGA/RTL resume](airpoint-fpga-rtl-resume.md) — ㈜에어포인트 FPGA/RTL; resume+portfolio Revised_JD_Fit 2026-07-13; fit ~74/~88; JD#1–4 mapped
- [Taeha controller SW resume](taeha-controller-sw-resume.md) — (주)태하 제어기 개발(SW) JD fit 94%; Final DOCX/PDF under `E:\이력서\태하 펌웨어 개발\`
- [Freelance engineering services](freelance-engineering-services.md) — AI freelancer plan: docs, RCA, resume, PLM decks, Streamlit tools; hub `C:\Users\namma\freelance\`
- [ChatGPT export index](chatgpt-export-index.md) — 41 conversations from 2026-07-02 export
- [ChatGPT library index](chatgpt-library-index.md) — 1141 library/knowledge-store files catalog
- [ChatGPT export preferences](chatgpt-export-preferences.md) — workflow themes and settings from export snapshot
- [ChatGPT stored memories](chatgpt-stored-memories.md) — identity, career, projects, preferences from ChatGPT cloud Memory export (2026-06-22)

- [Office docs via COM](office-docs-com-automation.md) — no Python/Node/LibreOffice here; edit pptx/docx via PowerShell + Office COM (with gotchas)
- [AVEVA PLM application deck](aveva-plm-application-deck.md) — English exec deck for AVEVA Marine PLM SME role, target HD Hyundai, in D:\이력서
- [AVEVA PLM FastAPI control plane](aveva-plm-fastapi-control-plane.md) — FastAPI reference API on E:\이력서\…\AVEVA_Marine_PLM_Control_Plane_…
- [ITT Cannon FAE resume](itt-cannon-fae-resume.md) — English resume for ITT Cannon Connector Korea FAE (via Adecco), D:\이력서\ITT Cannon
- [Genohco 시스템 체계 resume](genohco-system-resume.md) — 제노코 지원 이력서/포트폴리오 (rev2), D:\이력서\제노코 — 제노코 실직무=방산 정비장비(K2조준경·UAV) PL, 위성/TVAC 아님
- [ChatGPT / OpenAI preferences](chatgpt-preferences.md) — durable cross-platform preferences mirrored for ChatGPT cloud Memory
- [Cursor Agent OS](cursor-agent-os.md) — Claude Code–parity prompt programming stack (AGENTS.md, skills, rules, subagents, commands)
- [Claude parity evolution](claude-parity-evolution.md) — Claude strengths → Cursor actions; 30-day roadmap; default STAGES workflow
- [Prompting techniques](prompting-techniques.md) — 10 techniques (zero-shot→meta) as Cursor agent self-prompting; `/prompt-better` (2026-07-17)
- [Claude memory absorption](claude-memory-absorption.md) — Claude AGENTS + hub MCP + boot sequence for Cursor
- [Claude stored memories](claude-stored-memories.md) — Claude cloud Memory export (2026-07-11): PLM project, career, interview prefs
- [Claude export index](claude-export-index.md) — 27 Claude conversations from 2026-07-11 export
- [Claude Cowork parity](claude-cowork-parity.md) — Cowork scheduled tasks → Cursor Automations mapping
- [Career pipeline](career-pipeline.md) — active job applications (Claude joonha-hub MCP JSON)
- [Resume positioning strategy](resume-positioning-strategy.md) — one role axis per application; customer-facing value compression; 2–3p resume (2026-07-17)
- [LinkedIn automation activation](linkedin-automation-activation.md) — daily 09:00 cloud + 09:35 local stage-only pipeline

## Workspace tooling (added 2026-06-06)

**Cursor Automations (2026-07-05)** — primary scheduler (Glass; save drafts)
- Memory Sync Hourly (`5 * * * *`), Daily (`0 8 * * *`), Manual (webhook)
- LinkedIn Marine PLM Post (`0 9 * * 1-6`) — stage only, no auto-Post
- Specs: `.cursor\automations\README.md`
- Codex backup: `memory-sync-hourly`, `memory-sync-daily` (ACTIVE); `manual-claude-codex-sync` (PAUSED)

**MCP secrets (2026-07-05)** — tokens in `.cursor\mcp-secrets.env` (gitignored); launchers `start-notion-mcp.ps1`, `start-figma-mcp.ps1`

**Cursor Agent OS (2026-07-05)** — Claude Code–parity prompt programming
- Master: `C:\Users\namma\.cursor\AGENTS.md` (+ root copy `C:\Users\namma\AGENTS.md` 2026-07-13)
- Skills: `prompt-orchestrator`, `office-com-doc-qa`, `jd-resume-pipeline`, `memory-hub`, `claude-parity` in `.cursor\skills\`
- `resume-builder-com`, `plm-slide-builder` in `.cursor\skills\`
- Builders restored: 184 PS scripts in `.cursor\builders\` (git history 2026-07-05)
- Rules: `core-agent-behavior`, `claude-parity-workflow`, `office-com-windows`, `resume-jd-workflow` in `.cursor\rules\`
- User Rule: `Claude-parity Cursor Agent OS` (Settings → Rules)
- Commands: `/memory-sync`, `/jd-match`, `/resume-revise`, `/orchestrate`, `/claude-evolve`, `/plm-qa`
- Hooks: sessionStart + stop → memory sync only (durable capture via rules; stop followup removed 2026-07-13 — loop fix)
- Google AI Overview hardening applied 2026-07-13 — see `claude-parity-evolution.md`

**Codex skill `office-com-doc-qa`** — `C:\Users\namma\.codex\skills\office-com-doc-qa\`
- Trigger: Office COM으로 이력서/제안서 편집 + PDF/PNG QA 요청 시 `$office-com-doc-qa`로 호출
- Files: `SKILL.md`, `scripts\Export-OfficePdfPng.ps1` (DOCX→PDF→PNG, PPTX→PDF→PNG), `references\office-com-patterns.md`
- Requirement: **PowerShell 5.1** (`powershell.exe`) 필수; PS7에서 실행 시 조기 실패 메시지 출력
- Smoke-tested: DOCX→PDF→PNG ✓, PPTX→PDF→PNG ✓ (2026-06-06)

**Git auto-init** — `C:\Users\namma\.claude\auto_git_projects_entry.ps1`
- Auto-init is for new project folders under `C:\Users\namma\.claude\projects\C--Users-namma--claude`; as of 2026-06-21, `itt_work` and `plm_slide_work` are plain work folders with no local `.git` directory.
- `C:\Users\namma\.claude\projects\C--Users-namma--claude` 아래 새 폴더 생성 시 자동 `git init` helper. 2026-06-24 policy update: Windows login auto-start was disabled with the sync watcher; do not restore the `ClaudeProjectGitInitWatcher` Run key unless the user explicitly asks.
- 로그: `C:\Users\namma\.claude\cache\git-auto-init\watcher.log`
- 원격 origin은 미설정 (로컬 추적만)

**Git autosync** — `C:\Users\namma\.claude\git_autosync.ps1`
- `C:\Users\namma\.claude`, `C:\Users\namma\.codex\skills`, `C:\Users\namma\.codex\automations`, Codex `sessions`, `.cursor\memory`, root `.gitignore`/`.gitattributes` 변경 감지 후 `C:\Users\namma` 루트 repo에 자동 commit + push
- 2026-06-24 chat sync expansion: Codex chat/code/cowork session JSONL and imported-session metadata are tracked in the root repo.
- Root repo/origin guard: `C:\Users\namma\.claude\ensure_git_autosync_connection.ps1`
- 런처: `C:\Users\namma\.claude\git_autosync_launcher.ps1`
- 2026-06-24 policy update: Windows login auto-start was disabled at user request. Run sync manually from Codex Automations or with `Get-Content C:\Users\namma\.claude\start_claude_codex_sync.ps1 -Raw | Invoke-Expression`.
- 로그: `C:\Users\namma\.claude\cache\git-autosync\autosync.log`

**Cross-platform memory sync** — `C:\Users\namma\.cursor\bin\memory-sync.ps1`
- Hub: `C:\Users\namma\.cursor\memory` (Cursor MCP `memory-sync`)
- Mirrors: Claude memory, `.codex\memory`, ChatGPT local mirror at `.cursor\memory\platforms\chatgpt`
- Manual: `workspace memory sync`, Claude Cowork **Scheduled → memory-sync-manual** (Run), or Cursor Automation webhook
- Auto: Cursor `sessionStart` hook (synchronous — completes before session continues; check Hooks output channel or `sync.log`) + Windows logon + Codex schedules
- Logon install: `workspace memory logon install` (Task Scheduler; falls back to HKCU Run key)

**Codex automation `claude-project-monitor`** — 6시간마다 `C:\Users\namma\.claude` 전체 감시

**LinkedIn image QA** — 2026-06-24
- LinkedIn reference images must verify that every visible text label stays inside its backing shape/panel/chip.
- Upload LinkedIn images directly from the run's workspace outputs folder (`C:\Users\namma\Documents\Codex\YYYY-MM-DD\<topic-slug>\outputs`).
- Avoid connector/leader lines unless the line clearly terminates on the exact object or feature named by the label.
- Preferred LinkedIn image style is a detailed but compressed engineering infographic.
