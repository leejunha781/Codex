---
name: resume-jd-match
description: JD analysis, resume tailoring, match scoring, and safe document output for job applications. Use when user mentions JD, job description, resume, cover letter, or match score.
---

# Resume & JD Match Workflow

## When to Use

- User provides a job description (JD) or asks for resume tailoring
- User wants match score or shortlist analysis
- User needs Codex/Cursor prompt documents for application pipeline

## Prerequisites

- Read `MEMORY.md` for career facts — never invent experience
- Identify canonical resume/cover source files before editing
- Confirm language: Korean for analysis, English for outward documents

## Steps

### 1. JD Analysis (Korean OK)

Extract and structure:

- Role title, company, location
- Must-have vs nice-to-have requirements
- Keywords for ATS
- Gap analysis vs user's real experience (from MEMORY.md)
- Honest match score (%) with evidence per requirement

### 2. Resume Tailoring

- Map each JD requirement to a bullet with quantified outcome where possible
- Use action verbs; lead with impact
- English only for resume/cover unless user requests Korean
- Align with target role: FAE, Technical PM/BDM, or PM

### 3. Safe File Output

```
1. Locate canonical file (e.g. *_Resume_Final.docx)
2. Copy to new name: *_Targeted_YYYYMMDD_HHMMSS.docx or *_Revised_JD_Fit.docx
3. Edit ONLY the copy
4. Report: original path (preserved) + new output path
```

### 4. Match Scoring Output

Produce:

- Overall match %
- Requirement-by-requirement table (Met / Partial / Gap)
- Recommended talking points for interview
- Optional: `_Match_and_Codex_Cursor_Prompt.docx` content outline

### 5. Multi-Model Handoff (Optional)

If user wants Claude polish:

- Export draft bullets and gap narrative
- Suggest Claude Project prompt for second-pass tone review
- Cursor applies final edits to files

## Quality Checks

- [ ] No fabricated employers, dates, or certifications
- [ ] Canonical files untouched
- [ ] Match score justified with evidence
- [ ] Outward content in English
- [ ] Korean summary provided if user wrote in Korean

## Example Trigger Phrases

- "이 JD에 맞게 이력서 수정해줘"
- "Match this job description"
- "90% match shortlist"
- "Cover letter for [company]"
