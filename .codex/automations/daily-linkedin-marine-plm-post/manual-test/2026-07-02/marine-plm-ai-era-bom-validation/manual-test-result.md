# Manual test run — 2026-07-02 (dry-run, Linux cloud agent)

## Topic
- **slug:** marine-plm-ai-era-bom-validation
- **angle:** AI-era developer leadership — AI drafts PLM scripts, lead developer owns BOM gate and shipyard adoption proof
- **marine execution context:** BOM maturity, design-change control, CAD/ERP/MES interfaces, shipyard adoption gates

## Outputs (workspace dry-run; Windows live path shown for reference)

| Artifact | Workspace path | Windows live path |
|----------|----------------|-------------------|
| Post | `.codex/automations/daily-linkedin-marine-plm-post/manual-test/2026-07-02/marine-plm-ai-era-bom-validation/linkedin-post.md` | `C:\Users\namma\Documents\Codex\2026-07-02\marine-plm-ai-era-bom-validation\linkedin-post.md` |
| Image | `.codex/automations/daily-linkedin-marine-plm-post/manual-test/2026-07-02/marine-plm-ai-era-bom-validation/marine-plm-ai-era-bom-validation-infographic.png` | `C:\Users\namma\Documents\Codex\2026-07-02\marine-plm-ai-era-bom-validation\marine-plm-ai-era-bom-validation-infographic.png` |

## Image QA
- text-fit: **pass**
- leader-lines: **pass** (no ambiguous connectors)
- overlap QA: **pass** (panels, thumbnails, checklist strip clear)

## LinkedIn status
- **blocked** — Linux cloud agent cannot access LinkedIn Windows app or `C:\Users\namma\` paths
- On Windows: run automation manually once via Codex thread; composer should stop at "ready for final posting" without user confirmation

## Validation
- `automation.toml` prompt validation: **pass** (all required sections present, ACTIVE)
- `memory.md` template: **pass** (rotation log, QA template, prior history preserved)

## Windows sync command

```powershell
cd <repo-root>\.codex\automations
.\sync-daily-linkedin-automation.ps1
.\validate-daily-linkedin-automation.ps1
```
