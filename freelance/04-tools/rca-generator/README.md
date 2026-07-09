# RCA Report Generator

Streamlit demo for **Service 1: Engineering Document Automation**.

## Setup

```powershell
cd C:\Users\namma\freelance\04-tools\rca-generator
pip install -r requirements.txt
streamlit run app.py
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
