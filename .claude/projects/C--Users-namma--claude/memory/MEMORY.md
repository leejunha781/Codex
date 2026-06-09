# Memory Index

- [Office docs via COM](office-docs-com-automation.md) — no Python/Node/LibreOffice here; edit pptx/docx via PowerShell + Office COM (with gotchas)
- [AVEVA PLM application deck](aveva-plm-application-deck.md) — English exec deck for AVEVA Marine PLM SME role, target HD Hyundai, in D:\이력서
- [ITT Cannon FAE resume](itt-cannon-fae-resume.md) — English resume for ITT Cannon Connector Korea FAE (via Adecco), D:\이력서\ITT Cannon
- [MoaComm Korea HW resume](moacomm-korea-hw-resume.md) — Korean resume for 모아컴코리아 HW 설계 부장급, canonical ..._최종_v2.docx in D:\이력서\모아컴코리아 - HW 개발

## Workspace tooling (added 2026-06-06)

**Codex skill `office-com-doc-qa`** — `C:\Users\namma\.codex\skills\office-com-doc-qa\`
- Trigger: Office COM으로 이력서/제안서 편집 + PDF/PNG QA 요청 시 `$office-com-doc-qa`로 호출
- Files: `SKILL.md`, `scripts\Export-OfficePdfPng.ps1` (DOCX→PDF→PNG, PPTX→PDF→PNG), `references\office-com-patterns.md`
- Requirement: **PowerShell 5.1** (`powershell.exe`) 필수; PS7에서 실행 시 조기 실패 메시지 출력
- Smoke-tested: DOCX→PDF→PNG ✓, PPTX→PDF→PNG ✓ (2026-06-06)

**Git auto-init** — `C:\Users\namma\.claude\auto_git_projects_entry.ps1`
- `itt_work`, `plm_slide_work`, 기존 채팅 폴더 3개 로컬 git 초기화 완료 (`main` 브랜치, `.gitignore` 포함)
- `C:\Users\namma\.claude\projects\C--Users-namma--claude` 아래 새 폴더 생성 시 자동 `git init` — 레지스트리 Run 키로 로그인 시 백그라운드 감시 스크립트 자동 시작
- 로그: `C:\Users\namma\.claude\cache\git-auto-init\watcher.log`
- 원격 origin은 미설정 (로컬 추적만)

**Git autosync** — `C:\Users\namma\.claude\git_autosync.ps1`
- `C:\Users\namma\.claude`, `C:\Users\namma\.codex\skills`, `C:\Users\namma\.codex\automations` 변경 감지 후 `C:\Users\namma` 루트 repo에 자동 commit + push
- 런처: `C:\Users\namma\.claude\git_autosync_launcher.ps1` (중복 실행 방지 pid 파일 사용)
- 로그인 자동 시작: HKCU Run 키 `ClaudeCodexGitAutosync`
- 로그: `C:\Users\namma\.claude\cache\git-autosync\autosync.log`

**Codex automation `claude-project-monitor`** — 6시간마다 `C:\Users\namma\.claude` 전체 감시
- 확인 항목: 최근 변경, 빠진 핵심 파일, 메모리 불일치, 위험한 설정/문서 변경, `itt_work`·`plm_slide_work` 이슈
- 자동화 메모리: `C:\Users\namma\.codex\automations\claude-project-monitor\memory.md`
