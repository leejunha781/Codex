# Freelance Engineering Services — Joonha Lee

**Positioning:** Engineering documentation + RCA reports + PLM/digital thread planning + simple AI automation tools.

**Not competing as:** generic web developer or low-cost coding freelancer.

**Competing as:** shipbuilding · defense · satellite · system integration specialist who turns field issues and engineering requirements into customer-facing deliverables.

## Start order (ChatGPT plan)

| Priority | Service | Status |
|----------|---------|--------|
| 1 | Technical Documentation / RCA / Proposals | `01-services/01-technical-documentation.md` |
| 2 | Resume / LinkedIn / Global application docs | `01-services/02-resume-linkedin-proposals.md` |
| 3 | PLM / Digital Thread planning support | `01-services/03-plm-digital-thread.md` |
| 4 | Streamlit engineering automation tools | `04-tools/rca-generator/` |

## 30-day execution plan

| Week | Focus | Deliverables |
|------|-------|--------------|
| 1 | Define 3 core services | `01-services/*.md` |
| 2 | Portfolio samples | `02-samples/` (RCA, FAT/SAT, PLM deck) |
| 3 | Go-to-market | `03-marketing/` (LinkedIn, Fiverr, Upwork, outreach) |
| 4 | Demo tool | `04-tools/rca-generator/` Streamlit app |

## Folder map

```
freelance/
├── 01-services/          # Service definitions, pricing, scope
├── 02-samples/           # Portfolio samples (anonymized)
├── 03-marketing/         # Platform listings, outreach templates
├── 04-tools/             # Streamlit apps and automation
│   └── rca-generator/
└── templates/            # Reusable document templates
```

## Existing workspace assets to reuse

| Asset | Location | Use for |
|-------|----------|---------|
| JD → Resume pipeline | `.cursor/skills/jd-resume-pipeline/` | Service #2 |
| PLM slide builders | `.cursor/builders/plm_slide_work/` | Service #3 samples |
| Office COM doc QA | `.codex/skills/office-com-doc-qa/` | DOCX/PDF export |
| LinkedIn post staging | `Documents/Codex/` | Service marketing |
| AVEVA PLM deck memory | `.cursor/memory/aveva-plm-application-deck.md` | PLM sample content |

## Run RCA Generator (Week 4)

```powershell
cd C:\Users\namma\freelance\04-tools\rca-generator
pip install -r requirements.txt
streamlit run app.py
```

## Contact

- LinkedIn: https://www.linkedin.com/in/joonha-lee-20b518316/
- Email: leejunha781@gmail.com
