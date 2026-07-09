---
name: freelance-engineering-services
description: Run the AI freelancer plan for Joonha Lee — engineering documentation, RCA, resume/LinkedIn, PLM decks, Streamlit tools. Use when user asks about freelance work, client deliverables, Fiverr/Upwork gigs, or the 30-day execution plan.
---

# Freelance Engineering Services

Read `C:\Users\namma\.cursor\memory\freelance-engineering-services.md` and hub `C:\Users\namma\freelance\README.md`.

## Three core services (start here)

| # | Service | Definition | Skill to pair |
|---|---------|------------|---------------|
| 1 | Technical documentation / RCA / proposals | `freelance/01-services/01-technical-documentation.md` | RCA Generator app |
| 2 | Resume / LinkedIn / global application docs | `freelance/01-services/02-resume-linkedin-proposals.md` | `jd-resume-pipeline` |
| 3 | PLM / Digital Thread planning | `freelance/01-services/03-plm-digital-thread.md` | `plm-slide-builder` |

## Client delivery workflow (Service 1 — RCA example)

1. Intake raw notes from client (equipment, symptom, environment, tests, actions)
2. Open RCA Generator or use `freelance/templates/RCA_Report_Template.md`
3. Generate draft → apply engineering judgment (do not ship raw AI output)
4. Export DOCX; for final client PDF on Windows use `office-com-doc-qa` if COM polish needed
5. Deliver with document ID and verification section

## Portfolio samples

- `freelance/02-samples/RCA_Report_Sample_Satellite_Terminal_Link_Loss.md`
- `freelance/02-samples/FAT_SAT_Checklist_Sample_LEO_Terminal.md`
- `freelance/02-samples/PLM_Concept_Deck_Outline_Sample.md`

## Marketing

- `freelance/03-marketing/platform-listings-and-outreach.md`
- LinkedIn: stage only; never auto-Post without user confirmation

## RCA Generator (demo tool)

```powershell
cd C:\Users\namma\freelance\04-tools\rca-generator
pip install -r requirements.txt
streamlit run app.py
```

## Pricing

See each service file in `01-services/` for starter USD ranges. Adjust after first paid projects.

## Language

- User chat Korean → respond Korean
- Client deliverables, Fiverr/Upwork, LinkedIn → English unless asked otherwise

## Do not

- Position as cheap general web development
- Sell full PLM implementation on day one (sell concept decks and planning docs)
- Overwrite canonical resume templates in `D:\이력서\`
