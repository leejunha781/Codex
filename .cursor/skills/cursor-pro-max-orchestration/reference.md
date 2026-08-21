# Cursor Pro Max — Reference

## Model Map (Authoritative)

| Task | Preferred | Mode |
|------|-----------|------|
| Agent coding, refactors, PRs | GPT-5.3 Codex High / Composer 2.5 | Max for large repos |
| Architecture, specs, firmware RCA | Claude Opus or Sonnet | Max |
| Resume / JD / long docs | Claude Sonnet or Opus | Max |
| Fast small edits | Auto / GPT Mini | Default |
| Figma UI | Claude Sonnet + Figma MCP | — |
| Cloud Agents | Same as task | Dashboard |

## Project Skill Catalog

| Skill | Path | Auto-discover |
|-------|------|---------------|
| cursor-pro-max-orchestration | `.cursor/skills/cursor-pro-max-orchestration/` | Yes (omit disable) |
| embedded-codex-workflow | `.cursor/skills/embedded-codex-workflow/` | Yes |
| stm32-device-drivers | `.cursor/skills/stm32-device-drivers/` | Yes |
| taeha-dispense-controller | `.cursor/skills/taeha-dispense-controller/` | Yes |
| resume-jd-match | `.cursor/skills/resume-jd-match/` | Yes |
| 30-resume-jd-workflow | `.cursor/skills/30-resume-jd-workflow/` | Yes |
| career-interview-prep | `.cursor/skills/career-interview-prep/` | Yes |
| plm-systems-consulting | `.cursor/skills/plm-systems-consulting/` | Yes |

Invoke with `@skill-name` or `@.cursor/skills/<name>/SKILL.md`.

## Personal Mirror

For cross-repo use, mirrors may live under `~/.cursor/skills/` (same names). Prefer the **project** copy when working inside `leejunha781/Codex` so git stays source of truth.

## Cursor Settings Checklist (User)

1. Enable Claude Opus/Sonnet, GPT-5.3 Codex, GPT-5.4  
2. Max Mode for complex agent sessions  
3. Default Agent Model: Auto; switch to Codex/Claude for heavy work  
4. MCP: Notion, Linear (+ Figma if design)  
5. Cloud Agents linked to this repo  
6. Privacy Mode if handling defence/marine-sensitive material  

## Cost Discipline

- First-party Auto/Composer for routine edits  
- One Max-depth pass for design; one Codex pass for implement  
- Monitor [cursor.com/settings](https://cursor.com/settings) usage  

## Learning Loop

After each substantive task:

1. New durable fact? → `MEMORY.md`  
2. New technical pattern? → domain skill `reference.md`  
3. New workflow? → this skill or `embedded-codex-workflow`
