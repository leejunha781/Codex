# NAVARINO Account Manager — Revised JD Fit (2026-07-21)

## Goal
Tailor Joonha Lee’s CV for **Navarino Account Manager (South Korea, remote + customer visits)** with formal English, LinkedIn/MEMORY-aligned facts, and honest match uplift.

## Source paths (canonical — preserved, not overwritten)
- `e:\이력서\Account Manager\Job Ad_Account Manager.docx`
- `e:\이력서\Account Manager\Joonha_Lee_NAVARINO_Account_Manager_Professional_CV.docx`
- `e:\이력서\Account Manager\Navarino_한국_Account_Manager_회사조사_JD매칭_합격전략_20260719.docx`

## Outputs
| Artifact | Path |
|----------|------|
| Revised CV (EN) | `Joonha_Lee_NAVARINO_Account_Manager_Professional_CV_Revised_JD_Fit_20260721.docx` |
| Match / strategy (KO+EN) | `Navarino_한국_Account_Manager_회사조사_JD매칭_합격전략_20260721_Revised.docx` |
| Builder | `build_navarino_am_cv.py` |
| Local Word COM copy + PDF QA | `Copy-ToLocalAndWordComQa.ps1` |
| Cloud PDF/PNG QA renders | `qa-renders/` |

## Match score (honest)
- Before reframing: **~58%**
- After Revised JD Fit: **~74%**

## Visual QA (cloud proxy — LibreOffice)
- CV PDF pages: **2** (after layout tighten; was 3)
- Renders: `qa-renders/cv-page-1.png`, `cv-page-2.png`
- Checked: no sparse trailing page; no Genohco satellite/TVAC; PLM self-directed disclaimer present; Account Manager title clear

## Local Windows: copy beside originals + Word COM PDF QA
On the PC that has `E:\이력서` and Microsoft Word (PowerShell 5.1):

```powershell
# After git pull of this branch (or copy the folder):
cd <repo>\career\navarino-account-manager
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Copy-ToLocalAndWordComQa.ps1 `
  -SourceDir (Get-Location).Path `
  -DestDir "E:\이력서\Account Manager"
```

The script:
1. Confirms canonical originals exist and **does not overwrite them**
2. Copies revised DOCX next to them (new filenames only)
3. Exports PDF via **Word COM** (`SaveAs` format 17)
4. Renders PNG via `pdftoppm` or Ghostscript if available
5. Writes `E:\이력서\Account Manager\qa-renders\navarino-am-20260721\QA_NOTE.txt`

## Rebuild DOCX in cloud/repo
```bash
python3 build_navarino_am_cv.py
```

## Role axis
**A — Sales / FAE / Solution (Account Manager)**. Marine PLM is framed as self-directed digital-ops fluency only (AVEVA application prep), not as employed Marine PLM delivery.
