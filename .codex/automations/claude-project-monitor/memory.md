# Claude project monitor memory

## 2026-06-22 run
- Run time: 2026-06-22 09:04:37 +09:00 (Asia/Seoul local shell time).
- Read C:\Users\namma\.claude\AGENTS.md, project MEMORY.md, and project memory files for Office COM, ITT Cannon, AVEVA PLM, and Genohco.
- Since last run 2026-06-21T23:00:31Z, only C:\Users\namma\.claude\cache\git-autosync\autosync.log changed under .claude; root repo C:\Users\namma was clean when checked.
- Key active artifacts exist: ITT Final_Integrated_v2 DOCX/PDF, AVEVA default + V18 PPTX, Genohco rev2 DOCX/PDF/PPTX/PDF.
- AVEVA default/V18 PPTX hashes differ, but package inspection showed both have 217 entries and 42 slides; only docProps/core.xml compressed size differs, so no content divergence found.
- plm_slide_work\final_materials_v18_deck_render has 42 PNGs, matching V18 memory. No ~$ temp files found in .claude work folders.
- Findings to keep watching: C:\Users\namma\.claude\config.toml uses 
px despite the workspace no-Node constraint and has nable_all_tools_auto_approval=true; C:\Users\namma\.claude\Scheduled\daily-linkedin-post\SKILL.md says to run bash/date; old stale lock file remains at D:\이력서\AVEVA - Marine Principal Technical Support & Consultant – PLM SME, Busan\Proposal\~.pptx; UUID project repos remain no-HEAD nested repos but autosync is currently pushing cleanly.
