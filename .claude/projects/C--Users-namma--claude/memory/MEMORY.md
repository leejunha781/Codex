# Memory Index

- [Office docs via COM](office-docs-com-automation.md) — no Python/Node/LibreOffice here; edit pptx/docx via PowerShell + Office COM (with gotchas)
- [AVEVA PLM application deck](aveva-plm-application-deck.md) — English exec deck for AVEVA Marine PLM SME role, target HD Hyundai, in D:\이력서
- [ITT Cannon FAE resume](itt-cannon-fae-resume.md) — English resume for ITT Cannon Connector Korea FAE (via Adecco), D:\이력서\ITT Cannon
- [Genohco 시스템 체계 resume](genohco-system-resume.md) — 제노코 지원 이력서/포트폴리오 (rev2), D:\이력서\제노코 — 제노코 실직무=방산 정비장비(K2조준경·UAV) PL, 위성/TVAC 아님

## Workspace tooling (added 2026-06-06)

**Codex skill `office-com-doc-qa`** — `C:\Users\namma\.codex\skills\office-com-doc-qa\`
- Trigger: Office COM으로 이력서/제안서 편집 + PDF/PNG QA 요청 시 `$office-com-doc-qa`로 호출
- Files: `SKILL.md`, `scripts\Export-OfficePdfPng.ps1` (DOCX→PDF→PNG, PPTX→PDF→PNG), `references\office-com-patterns.md`
- Requirement: **PowerShell 5.1** (`powershell.exe`) 필수; PS7에서 실행 시 조기 실패 메시지 출력
- Smoke-tested: DOCX→PDF→PNG ✓, PPTX→PDF→PNG ✓ (2026-06-06)

**Git auto-init** — `C:\Users\namma\.claude\auto_git_projects_entry.ps1`
- Auto-init is for new project folders under `C:\Users\namma\.claude\projects\C--Users-namma--claude`; as of 2026-06-21, `itt_work` and `plm_slide_work` are plain work folders with no local `.git` directory.
- `C:\Users\namma\.claude\projects\C--Users-namma--claude` 아래 새 폴더 생성 시 자동 `git init` — 레지스트리 Run 키로 로그인 시 백그라운드 감시 스크립트 자동 시작
- 로그: `C:\Users\namma\.claude\cache\git-auto-init\watcher.log`
- 원격 origin은 미설정 (로컬 추적만)

**Git autosync** — `C:\Users\namma\.claude\git_autosync.ps1`
- `C:\Users\namma\.claude`, `C:\Users\namma\.codex\skills`, `C:\Users\namma\.codex\automations`, Codex `sessions`, `archived_sessions`, pasted-text `attachments`, `session_index.jsonl`, `external_agent_session_imports.json`, `process_manager\chat_processes.json`, root `.gitignore`/`.gitattributes` 변경 감지 후 `C:\Users\namma` 루트 repo에 자동 commit + push
- 2026-06-21 hardening: autosync now detects nested `.git` boundaries under watched paths and excludes those paths before root `git add`, so no-HEAD nested repos do not cause repeated `git add failed` loops. The UUID project-folder skip remains, but generic nested-repo filtering is now the main guard.
- 2026-06-24 chat sync expansion: Codex chat/code/cowork session JSONL and imported-session metadata are tracked in the root repo; Claude UUID chat helper folders still skip nested repo metadata but allow `subagents/*.json[l]` and `tool-results/*.txt`.
- Root repo/origin guard: `C:\Users\namma\.claude\ensure_git_autosync_connection.ps1` verifies `C:\Users\namma\.git` and origin `https://github.com/leejunha781/Codex.git`, then starts autosync.
- 런처: `C:\Users\namma\.claude\git_autosync_launcher.ps1` (중복 실행 방지 pid 파일 사용)
- 로그인 자동 시작: HKCU Run 키 `ClaudeCodexGitAutosync` → `ensure_git_autosync_connection.ps1`
- 로그: `C:\Users\namma\.claude\cache\git-autosync\autosync.log`

**Codex automation `claude-project-monitor`** — 6시간마다 `C:\Users\namma\.claude` 전체 감시
- 확인 항목: 최근 변경, 빠진 핵심 파일, 메모리 불일치, 위험한 설정/문서 변경, `itt_work`·`plm_slide_work` 이슈
- 자동화 메모리: `C:\Users\namma\.codex\automations\claude-project-monitor\memory.md`

**LinkedIn image QA** — 2026-06-24
- LinkedIn reference images must verify that every visible text label stays inside its backing shape/panel/chip. Measure text before rendering where possible; otherwise enlarge the backing shape, shorten/wrap text, or reduce font size. Do not post images where text crosses outside the intended background shape.
- Upload LinkedIn images directly from the run's workspace outputs folder (`C:\Users\namma\Documents\Codex\YYYY-MM-DD\<topic-slug>\outputs`). Do not create separate direct-upload copies in generated-images folders unless the user explicitly asks.
- Avoid connector/leader lines unless the line clearly terminates on the exact object or feature named by the label. If target mapping is ambiguous, omit the line and place the label near the visual element instead.
- Preferred LinkedIn image style is a detailed but compressed engineering infographic, like a solution overview poster: strong title/subtitle, 3-4 numbered blocks, compact bullets, central architecture/system visual, checklist/takeaway strip, clean icons, and immediately readable English labels.
- Do not make LinkedIn images feel flat or overly schematic. Prefer a natural, premium, non-AI-looking maritime/shipyard/operations-room visual foundation with cinematic but believable lighting, rich blue/teal/amber contrast, polished dashboard depth, and vivid but restrained highlights behind the compressed infographic content.
