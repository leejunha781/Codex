2026-06-06 run (~11m)

- Read `AGENTS.md`, `projects/C--Users-namma--claude/memory/MEMORY.md`, and both project memory files.
- Workspace is not a git repo; monitoring is file/activity based.
- Canonical AVEVA deck exists at `D:\이력서\AVEVA - Marine Principal Technical Support & Consultant – PLM SME, Busan\Proposal\Future_Industrial_PLM_Meeting_Deck_EN_Hybrid_Safety_Final_v4.pptx`, but disk metadata moved past project memory: file now `27,688,693` bytes and `2026-06-06 13:16:24`, while memory still says `25,638,360` bytes and `12:52:43`.
- `plm_slide_work` QA folders are mostly healthy but stale relative to the canonical deck: `visualreview_current\audit.txt` and `meeting_current\audit.txt` still target 33-slide `...Rebuilt_v2...` artifacts, not the 34-slide canonical v4 deck.
- `itt_work` contains many intermediate backups/logs. The canonical resume exists at `D:\이력서\ITT Cannon\Joonha_Lee_ITT_Cannon_FAE_Resume_Final_Integrated_v2.docx` (`2026-06-05 22:32:09`).
- `itt_work\dump_full.txt` still reflects a stale/inconsistent structure: duplicate Daeyang detail tables and no visible Intellian detail table, so local diagnostic dumps are not trustworthy as current-state evidence.
- Resume memory already flags `IELTS 5.0` as an optional content risk; no new structural break proven in the canonical doc from this run.
- `AGENTS.md` path guidance is stale for this workspace: it points to `projects/C--Users-namma--Codex/memory/`, but the live memory path here is `projects/C--Users-namma--claude/memory/`.
- A titleless `POWERPNT` process was present during inspection, and COM attach failed with `0x80070520`; matches the recorded stale-window/orphan risk. Avoid saving any stale visible PowerPoint window until the deck is reopened from disk.
