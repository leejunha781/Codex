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
| Cursor Pro | Agent coding, git, MCP, Cloud Agents, rules/skills |
| Claude Pro | Long reasoning, document polish, Projects |
| Codex / GPT-5.3 | Implementation, refactors, automation |

## Embedded / Dispense (Durable)

| Fact | Detail |
|------|--------|
| KSS-III MCU | **NXP LPC1769** (Cortex-M3) firmware, DIO/AIO, RS-485/RS-232, ship install |
| Intellian STM32 | STM32F existing BLDC code analysis / validation / tuning (Conscan·Consearch) |
| Research | TI RM57L843 EtherCAT / servo / 25 kHz PWM — research, not production claim |
| Taeha focus | Precision dispense controller SW: recipe, calibration, motors, Modbus, safety |
| Skills | `cursor-pro-max-orchestration`, `embedded-codex-workflow`, `stm32-device-drivers`, `taeha-dispense-controller` |
| Notion hub | 태하 제어기 개발(SW) 기술면접 준비 허브 |
| Playbook | `docs/EMBEDDED_DEV_STACK.md` |

Accuracy: do not equate LPC1769 with STM32; do not invent Taeha private architecture.

### Multi-model default

Substantive work → Max Mode + Claude (design) + Codex (implement) + Cursor (verify/PR). Route via `@cursor-pro-max-orchestration`.

## Repo

- **Codex** (`leejunha781/Codex`) — AI configuration hub, rules, skills, career tooling

## Related Files

- `chatgpt-preferences.md` — exported ChatGPT preferences
- `docs/SETUP_CURSOR_PRO_CLAUDE_CODEX.md` — subscription & settings guide
- `.cursor/rules/` — project rules
- `.cursor/skills/` — workflow skills (career + embedded)
