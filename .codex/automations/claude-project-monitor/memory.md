2026-06-06 run (~9m)

- Read `AGENTS.md`, automation memory, and `projects\C--Users-namma--claude\memory\*.md`.
- AVEVA deck state drifted again: project memory still treats `D:\이력서\AVEVA - Marine Principal Technical Support & Consultant – PLM SME, Busan\Proposal\Future_Industrial_PLM_Meeting_Deck_EN.pptx` / `..._EN_V5.pptx` (37 slides, `27,704,970` bytes, `2026-06-06 15:48:08`) as canonical, but the newest same-family file on disk is `...\Future_Industrial_PLM_Meeting_Deck_EN__v4.pptx` (`32,903,742` bytes, `2026-06-06 16:03`), so "latest deck" is now ambiguous and needs explicit designation.
- `plm_slide_work` is active and healthy for the newest V5 merge pass: fresh QA render folders `merge_v4_images_final*` exist through `18:56`, but older audit artifacts `visualreview_current\audit.txt` and `meeting_current\audit.txt` still describe 33-slide `...Rebuilt_v2...` decks and should not be used as current evidence.
- `itt_work` looks stable and mostly archival; canonical resume files still exist in `D:\이력서\ITT Cannon\` with alias copies `Joonha_Lee_ITT_Cannon_FAE_Resume.docx/.pdf` saved minutes after `...Final_Integrated_v2...`, so there is mild filename ambiguity but no new break found.
- `AGENTS.md` remains stale for memory-path guidance: it still points to `projects/C--Users-namma--Codex/memory/`, while the live workspace path is `projects/C--Users-namma--claude/memory/`.
- Auto project tracking is inconsistent: watcher log says repos were initialized under `projects\C--Users-namma--claude\*`, but sampled project folders such as `35b0a916-8003-4154-8d9a-d79a2f0a1cff` and newer `6408bafd-8d68-49e5-8a93-5305b48a717d` have no `.git` directory. Treat auto-init as unreliable until verified/fixed.
- Git autosync appears healthy at the root-repo level: `cache\git-autosync\autosync.log` shows repeated successful pushes through `2026-06-06 19:14`, with another pending change detected at `19:16`.

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
