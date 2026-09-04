# MEMORY — Joonha Lee Context Hub

> Durable context for Cursor, Claude, and Codex agents. Update when preferences or facts change.

## Identity

| Field | Value |
|-------|-------|
| Name | Joonha Lee (이준하) |
| Location | Gunpo-si, Gyeonggi-do, Korea |
| LinkedIn | https://www.linkedin.com/in/joonha-lee-20b518316/ |

## Career Summary

- **21+ years** — marine, PLM, naval shipbuilding, defence electronics, satellite comms, system integration
- **Daeyang Electric (15 years)** — ROK Navy ship/submarine programmes
- **Intellian** — LEO satellite terminal integration/validation, FAT/SAT, commissioning
- **Education:** M.S. Embedded Systems (Pusan National Univ.), B.S. Control & Instrumentation (Univ. of Ulsan)
- **Target roles:** FAE, Technical PM/BDM, Project Management

## Communication Preferences

- Tone: calm, grateful, technically grounded, customer-oriented, long-term committed
- Interview answers: 30–60 sec, simple English, one concrete example
- Korean email: `검토해 주셔서 감사합니다.` / closing `감사합니다. 이준하 드림`
- Avoid: desperation, negative past-employer comments, "I do not know PLM"

## Resume & JD Workflow

1. JD analysis → resume tailoring → match scoring → prompt docs
2. Pipeline: Codex draft → Claude review → Cursor final
3. Output naming: `_Match_and_Codex_Cursor_Prompt.docx`, `_90pct_Actual_Match_JD_Shortlist_with_Links_`
4. **Never overwrite** canonical resume/cover templates — use `_Revised_JD_Fit`, `_Targeted_YYYYMMDD` suffixes
5. Korean for analysis prompts; English for outward resume/cover content

## AI Stack

| Service | Use For |
|---------|---------|
| Cursor Pro | Agent coding, git, MCP, Cloud Agents, rules/skills. Local Windows sandbox = WSL2 Landlock (not Cloud Agents). |
| Claude Pro | Long reasoning, document polish, Projects |
| Codex / GPT-5.3 | Implementation, refactors, automation |

## Repo

- **Codex** (`leejunha781/Codex`) — AI configuration hub, rules, skills, career tooling

## Related Files

- `chatgpt-preferences.md` — exported ChatGPT preferences
- `docs/SETUP_CURSOR_PRO_CLAUDE_CODEX.md` — subscription & settings guide
- `.cursor/rules/` — project rules
- `.cursor/skills/` — workflow skills
