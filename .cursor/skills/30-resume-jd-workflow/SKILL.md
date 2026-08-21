---
name: 30-resume-jd-workflow
description: Resume tailoring, JD matching, and document workflow — never overwrite canonical files
---

# Resume & JD Matching Workflow

Invoke via `@resume-jd-workflow` or when user mentions JD, resume, cover letter, or job match.

## Pipeline

1. **JD analysis** (Korean prompts OK) — requirements, keywords, gaps, match score.
2. **Resume tailoring** — English outward content; align bullets to JD evidence.
3. **Match scoring** — Honest % with shortlist rationale.
4. **Output artifacts** — `_Match_and_Codex_Cursor_Prompt.docx`, `_90pct_Actual_Match_JD_Shortlist_with_Links_` naming patterns.

## Multi-Model Flow

- **Cursor:** JD parse, draft bullets, file generation, repo organization.
- **Claude Pro:** Tone polish, gap narrative, cover letter refinement.
- **Codex:** Automation scripts for batch JD processing if needed.

## File Safety (Mandatory)

- **NEVER** overwrite canonical templates (`*_Resume_Final.docx`, `*_Template.docx`, `*_Cover_Letter_Final.docx`).
- Copy to new path with suffix: `_Revised_JD_Fit`, `_Targeted_YYYYMMDD`, `_backup_YYYYMMDD`.
- Report both paths: original (preserved) and new output.

## Skill

Load `.cursor/skills/resume-jd-match/SKILL.md` before executing this workflow.
