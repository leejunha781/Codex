2026-06-21 autosync fix follow-up

- `C:\Users\namma\.claude\git_autosync.ps1` was hardened after the monitor report below: watched paths are now filtered for nested `.git` boundaries before root `git add`, so no-HEAD nested repos are skipped by autosync instead of causing repeated `git add failed` loops. Verification used a throwaway no-HEAD nested repo under `.claude`; raw `git status` saw it, but the patched collector excluded it.

2026-06-21 project monitor follow-up (~9m)

- `C:\Users\namma\.claude\AGENTS.md` is now aligned with the live `.claude` workspace paths; the earlier stale `.Codex` / `projects/C--Users-namma--Codex` guidance is resolved.
- `D:\이력서` was mounted this run. Canonical docs verified on disk: `D:\이력서\ITT Cannon\Joonha_Lee_ITT_Cannon_FAE_Resume_Final_Integrated_v2.docx/.pdf` and the AVEVA deck alias set `D:\이력서\AVEVA - Marine Principal Technical Support & Consultant – PLM SME, Busan\Proposal\Future_Industrial_PLM_Meeting_Deck_EN.pptx`, `..._EN_V18.pptx`, and `..._EN_Final.pptx` all exist with 2026-06-09 timestamps around the recorded V18 state.
- `C:\Users\namma\.claude\cache\git-autosync\autosync.log` shows the autosync issue is no longer active after the 2026-06-21 11:13 restart. There were startup failures and one non-fast-forward push rejection at 11:09, but pushes have been succeeding continuously from 11:34 through 13:40, and root `git status` is clean.
- Auto-init remains partially inconsistent. `C:\Users\namma\.claude\cache\git-auto-init\watcher.log` still claims repos were initialized for `C:\Users\namma\.claude\projects\C--Users-namma--claude\6408bafd-8d68-49e5-8a93-5305b48a717d` and `...\baeb5d66-046c-4315-b5b7-4f26997e0fc2`, but both directories currently have no `.git` directory, while sibling project folders do.
- `C:\Users\namma\.claude\projects\C--Users-namma--claude\memory\MEMORY.md` is stale by omission relative to current activity. Today’s active local work areas `C:\Users\namma\.claude\abb_resume_work\`, `C:\Users\namma\.claude\hrbros_resume_work\`, and `C:\Users\namma\.claude\aveva_bdm_resume_work\` plus their canonical `D:\이력서\ABB - Portfolio and Industry Manager\` outputs are not indexed there yet.
- `C:\Users\namma\.claude\plm_slide_work\word_future_direction\~$EVA_Marine_Future_Direction_Meeting_Materials_EN.html` is still a stale Office lock/temp artifact. No active `POWERPNT` process was present during the check; only a visible `WINWORD` session for an ABB resume was open.

Run time: ~9 minutes.

2026-06-21 project monitor summary

- Root autosync is still blocked, and the failure mode is now fully confirmed. `C:\Users\namma\.claude\git_autosync.ps1` collects changed paths from the parent repo without excluding nested repos (`C:\Users\namma\.claude\git_autosync.ps1:49-57`, `60-87`). In `C:\Users\namma`, `git add --all --` fails on `C:\Users\namma\.claude\projects\C--Users-namma--claude\6408bafd-8d68-49e5-8a93-5305b48a717d\` and `...\baeb5d66-046c-4315-b5b7-4f26997e0fc2\` because both repos are initialized but have **no initial commit** (`git status` = `## No commits yet on main`; `git rev-parse HEAD` fails). The autosync loop kept retrying through `C:\Users\namma\.claude\cache\git-autosync\autosync.log:1758-1837`.
- `C:\Users\namma\.claude\AGENTS.md` is still stale versus the actual workspace. It still points to `C:\Users\namma\.Codex\itt_work\`, `C:\Users\namma\.Codex\plm_slide_work\`, and `projects/C--Users-namma--Codex/memory` (`C:\Users\namma\.claude\AGENTS.md:62-69`), while the live workspace and project memory are under `.claude`.
- Project memory now overstates git coverage for work folders. `C:\Users\namma\.claude\projects\C--Users-namma--claude\memory\MEMORY.md:16-17` says `itt_work` and `plm_slide_work` were git-initialized, but both folders currently have no `.git` directory at all.
- Canonical Office document locations could not be verified this run because `D:\이력서` is not mounted in this environment. The paths referenced in `AGENTS.md` and the resume/deck memories are currently unavailable, so checks were limited to the local working folders under `C:\Users\namma\.claude\`.
- `C:\Users\namma\.claude\plm_slide_work\fix_v6_slide13_14_readability.ps1` still hardcodes V6-era deck paths and backup naming (`:3-10`), so it no longer matches the later PLM deck lineage recorded in memory.
- A stale Office temp artifact still exists in the local PLM work area with no live Office process: `C:\Users\namma\.claude\plm_slide_work\word_future_direction\~$EVA_Marine_Future_Direction_Meeting_Materials_EN.html` (last written 2026-06-06 22:51). No `WINWORD` or `POWERPNT` process was running during this check.
- `C:\Users\namma\.claude\itt_work\` remains stable. The latest local scripts are still `rebuild_clean.ps1` and `render_all.ps1` from 2026-06-05, with no newer churn detected.
- Recent workspace churn since the last run is concentrated in `C:\Users\namma\.claude\genohco_system_work\` plus memory/index updates on 2026-06-17. No additional risky configuration changes were found beyond the existing autosync backlog.

Run time: ~14 minutes.

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

2026-06-06 run (~8m)

- Saved the user-supplied AVEVA software code package `D:\이력서\AVEVA - Marine Principal Technical Support & Consultant – PLM SME, Busan\Proposal\AVEVA_Marine_PLM_Control_Plane_FastAPI_SQL_IndustrialAI_FINAL_v2.zip`.
- Extracted it to `D:\이력서\AVEVA - Marine Principal Technical Support & Consultant – PLM SME, Busan\Proposal\AVEVA_Marine_PLM_Control_Plane_FastAPI_SQL_IndustrialAI_FINAL_v2\`.
- Recorded ZIP metadata: `185,527` bytes, LastWriteTime `2026-06-06 21:14:22`, SHA256 `B743D8482D3CC353C1059C5F1E29A48DEB22F255F60D6C7B7AF827160E7721A7`.
- Created `C:\Users\namma\.claude\plm_slide_work\software_code_final_v2_inventory.txt` with original-file hashes and `C:\Users\namma\.claude\plm_slide_work\software_code_final_v2_code_contents.txt` with the consolidated source/SQL/docs/OpenAPI/test text dump.
- Package scope checked: FastAPI app, PostgreSQL init SQL, OpenAPI JSON, docs, tests, Docker Compose, and requirements. Original ZIP has 108 files; extraction has 109 after adding `CODE_CONTENTS_SAVED.txt`.
- Did not run Python/Node/tests due workspace constraints; this run only saved, extracted, hashed, and indexed the code contents.
