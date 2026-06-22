# Claude project monitor memory

## 2026-06-22 run
- Run time: 2026-06-22 10:25 to 11:04:52 +09:00 (Asia/Seoul local shell time).
- Read C:\Users\namma\.claude\AGENTS.md, C:\Users\namma\.claude\projects\C--Users-namma--claude\memory\MEMORY.md, and tails/lists for the active project memories and Office work folders. No Office COM document open was needed; no WINWORD/POWERPNT process was running.
- Root repo C:\Users\namma was clean when checked. Autosync log shows successful pushes through 2026-06-22 10:04 +09:00. Latest since last-run git activity observed: C:\Users\namma\.codex\automations\claude-project-monitor\memory.md only.
- Active memories for ITT Cannon, AVEVA PLM, and Genohco still match existing canonical artifacts. ITT Final_Integrated_v2 DOCX/PDF exist; Genohco rev2 DOCX/PDF/PPTX/PDF exist; AVEVA PLM V18/default lineage is documented and C:\Users\namma\.claude\plm_slide_work\final_materials_v18_deck_render still contains 42 PNGs page-01.png through page-42.png.
- New actionable findings: C:\Users\namma\.claude\projects\C--Users-namma--claude\memory\MEMORY.md lacks entries for June 21 work folders C:\Users\namma\.claude\abb_resume_work, C:\Users\namma\.claude\hrbros_resume_work, and C:\Users\namma\.claude\aveva_bdm_resume_work, all of which contain build/extract scripts and rendered/document outputs in D:\이력서. Add memory entries/canonical artifact notes before future edits.
- Existing watch items remain: C:\Users\namma\.claude\config.toml was added 2026-06-22 and uses npx plus enable_all_tools_auto_approval=true despite workspace no-Node constraints; C:\Users\namma\.claude\Scheduled\daily-linkedin-post\SKILL.md was added/modified 2026-06-22 and still instructs a bash/date command.
- Cleanup candidates: stale hidden Office lock/temp files with no active Office process at C:\Users\namma\.claude\plm_slide_work\word_future_direction\~$EVA_Marine_Future_Direction_Meeting_Materials_EN.html, D:\이력서\에이치알브로스 - 무선통신 전문기업 통신장비 연구소장\~$하_통신장비_연구소장_사람인양식_이력서_자기소개서정정.docx, plus three older stale locks in the AVEVA PLM SME folder.
- Nested UUID project folders under C:\Users\namma\.claude\projects\C--Users-namma--claude still have .git directories but no initial commits/HEAD objects; several contain untracked .gitignore, subagents, or tool-results artifacts. Autosync is currently protected by nested-repo exclusion, but these repos remain a low-level hygiene risk.

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

## 2026-06-22 run
- Run time: 2026-06-22 12:02 to 2026-06-22 12:04:41 +09:00 (Asia/Seoul local shell time).
- Read C:\Users\namma\.claude\AGENTS.md, C:\Users\namma\.claude\projects\C--Users-namma--claude\memory\MEMORY.md, and the existing automation memory. Root repo C:\Users\namma was clean; latest post-last-run git activity was only C:\Users\namma\.codex\automations\claude-project-monitor\memory.md.
- Confirmed active Office baselines still exist: ITT Final_Integrated_v2 DOCX/PDF; AVEVA PLM default, V18, Final PPTX files; V18 render folder still has page-01.png through page-42.png. No WINWORD/POWERPNT process was running.
- Current actionable findings unchanged: C:\Users\namma\.claude\projects\C--Users-namma--claude\memory\MEMORY.md lacks entries for June 21 ABB, HRBros, and AVEVA BDM resume work folders; C:\Users\namma\.claude\config.toml uses npx and enable_all_tools_auto_approval=true; C:\Users\namma\.claude\Scheduled\daily-linkedin-post\SKILL.md still instructs bash/date; stale Office lock/temp files remain in HRBros, AVEVA PLM, and plm_slide_work; nested UUID project repos still contain untracked artifacts.
