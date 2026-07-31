# RCA Report Generator

Streamlit demo for **Service 1: Engineering Document Automation**.

## Setup

```powershell
cd C:\Users\namma\freelance\04-tools\rca-generator
python -m pip install -r requirements.txt
python -m streamlit run app.py
```

If PowerShell says `streamlit` is not recognized, use `python -m streamlit`
as shown above. The Python module is installed, but the `streamlit.exe`
launcher may not be on the Windows PATH.

Shortcut:

```powershell
.\run_app.cmd
```

## Features

- Structured RCA intake form (equipment, symptom, environment, tests, root cause, actions, verification)
- Live Markdown preview
- Export to `.md` and `.docx`
- **Load sample data** — pre-fills satellite terminal link-loss portfolio example

## Use in freelance workflow

1. Client sends raw notes → paste into form  
2. Generate draft → review with engineering judgment  
3. Export DOCX → final polish in Word (optional Office COM QA)

## Next tools (backlog)

- FAT/SAT checklist generator  
- Issue tracking dashboard (SQLite + CSV export)  
- Proposal scope / WBS generator
