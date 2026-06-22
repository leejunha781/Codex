# Claude project monitor memory

## 2026-06-22 run
- Run time: 2026-06-22 10:01:21 +09:00 to about 10:16 +09:00 (Asia/Seoul local shell time).
- Read C:\Users\namma\.claude\AGENTS.md, C:\Users\namma\.claude\projects\C--Users-namma--claude\memory\MEMORY.md, the Office COM QA skill, and the project memory files for Office COM, ITT Cannon, AVEVA PLM, and Genohco.
- Since last run 2026-06-22T00:01:02Z, notable .claude changes were C:\Users\namma\.claude\cache\git-autosync\autosync.log and auto-created nested .git directories under UUID project folders. Root repo C:\Users\namma was clean; latest autosync commit observed was 7e54c99 at 2026-06-22 09:05 +09:00.
- Key artifacts still exist: ITT Final_Integrated_v2 DOCX/PDF, Genohco rev2 DOCX/PDF/PPTX/PDF, AVEVA default + V18 PPTX. AVEVA default and V18 differ by hash/1 byte but package inspection showed both have 217 entries and only docProps/core.xml compressed size differs (429 vs 430), so no slide-content divergence found.
- plm_slide_work\final_materials_v18_deck_render still has 42 PNGs page-01.png through page-42.png, matching V18 memory. No WINWORD/POWERPNT processes were running during inspection.
- New/changed findings: stale temp file C:\Users\namma\.claude\plm_slide_work\word_future_direction\~$EVA_Marine_Future_Direction_Meeting_Materials_EN.html remains (162 bytes, last write 2026-06-06 22:51:21, no active Office process). UUID nested repos now have HEADs, but several contain untracked .gitignore/subagents/tool-results artifacts. Existing watch items remain: C:\Users\namma\.claude\config.toml uses npx and enable_all_tools_auto_approval=true; C:\Users\namma\.claude\Scheduled\daily-linkedin-post\SKILL.md says to run bash/date.

## 2026-06-22 run
- Run time: 2026-06-22 09:04:37 +09:00 (Asia/Seoul local shell time).
- Read C:\Users\namma\.claude\AGENTS.md, project MEMORY.md, and project memory files for Office COM, ITT Cannon, AVEVA PLM, and Genohco.
- Since last run 2026-06-21T23:00:31Z, only C:\Users\namma\.claude\cache\git-autosync\autosync.log changed under .claude; root repo C:\Users\namma was clean when checked.
- Key active artifacts exist: ITT Final_Integrated_v2 DOCX/PDF, AVEVA default + V18 PPTX, Genohco rev2 DOCX/PDF/PPTX/PDF.
- AVEVA default/V18 PPTX hashes differ, but package inspection showed both have 217 entries and 42 slides; only docProps/core.xml compressed size differs, so no content divergence found.
- plm_slide_work\final_materials_v18_deck_render has 42 PNGs, matching V18 memory. No ~$ temp files found in .claude work folders.
- Findings to keep watching: C:\Users\namma\.claude\config.toml uses 
px despite the workspace no-Node constraint and has nable_all_tools_auto_approval=true; C:\Users\namma\.claude\Scheduled\daily-linkedin-post\SKILL.md says to run bash/date; old stale lock file remains at D:\이력서\AVEVA - Marine Principal Technical Support & Consultant – PLM SME, Busan\Proposal\~.pptx; UUID project repos remain no-HEAD nested repos but autosync is currently pushing cleanly.
