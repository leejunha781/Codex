# claude-project-monitor memory

## 2026-06-22T13:05:03.7239082+09:00
- Read AGENTS.md and projects/C--Users-namma--claude/memory/MEMORY.md; automation memory was missing, so this file was created as the baseline.
- Checked root git status from C:\Users\namma: clean before writing this memory file. Autosync log showed healthy pushes through 2026-06-22 12:05 KST.
- Since last run (2026-06-22 12:02 KST), only C:\Users\namma\.claude\cache\git-autosync\autosync.log changed; no D:\이력서 files changed.
- Verified active Office areas: C:\Users\namma\.claude\itt_work quiet since 2026-06-05; D:\이력서\ITT Cannon canonical Final_Integrated_v2 docx/pdf present. AVEVA Proposal default/V18/final files present; V18 render folder has 42 PNG pages; no POWERPNT/WINWORD process running. Stale lock file remains at D:\이력서\AVEVA - Marine Principal Technical Support & Consultant – PLM SME, Busan\Proposal\~$Future_Industrial_PLM_Meeting_Deck_EN_V15.pptx.
- Findings to carry forward: (1) C:\Users\namma\.claude\config.toml registers windows-computer-use via npx, but node/npx are absent from PATH, so the MCP entry is likely broken. (2) C:\Users\namma\.claude\projects\C--Users-namma--claude\memory\office-docs-com-automation.md line 20 still recommends Set-ExecutionPolicy -Scope Process Bypass, conflicting with line 33 and AGENTS.md line 9.

## 2026-06-22T14:03:07.9889209+09:00
- Read C:\Users\namma\.claude\AGENTS.md, C:\Users\namma\.claude\projects\C--Users-namma--claude\memory\MEMORY.md, and prior automation memory.
- Checked recent changes since 2026-06-22T13:02:06+09:00: only C:\Users\namma\.claude\cache\git-autosync\autosync.log changed in workspace; no D:\이력서 files changed in that window.
- Root git status at C:\Users\namma was clean before this memory write; latest autosync log showed successful pushes through 2026-06-22 13:05 KST.
- Verified active Office areas: ITT canonical Final_Integrated_v2 docx/pdf present; AVEVA default and V18 PPTX present; V18 render folder has 42 PNG files; no WINWORD/POWERPNT processes running.
- Actionable findings: C:\Users\namma\.claude\config.toml still points windows-computer-use at npx while node/npm/npx are missing; C:\Users\namma\.claude\projects\C--Users-namma--claude\memory\office-docs-com-automation.md line 20 still recommends Set-ExecutionPolicy Bypass despite line 33/AGENTS.md saying it is blocked; identical ABB resume docx/pdf files are duplicated in D:\이력서\AVEVA - Business Development Manager, likely a filing inconsistency.
- Other hygiene: stale Office lock/temp files persist, including D:\이력서\에이치알브로스 - 무선통신 전문기업 통신장비 연구소장\~.docx and the older AVEVA V15 lock file.
