# Cursor Skills Index

Source of truth for Agent Skills in this repo. Invoke with `@skill-name` or `@.cursor/skills/<name>/SKILL.md`.

## Orchestration (load first on hard tasks)

| Skill | What |
|-------|------|
| [cursor-pro-max-orchestration](cursor-pro-max-orchestration/SKILL.md) | Cursor Pro Max + Claude + Codex on every substantive task |
| [embedded-codex-workflow](embedded-codex-workflow/SKILL.md) | Firmware vertical slice + Codex prompt template |

## Embedded / Dispense

| Skill | What |
|-------|------|
| [stm32-device-drivers](stm32-device-drivers/SKILL.md) | STM32 CubeMX/HAL/LL, ISR/DMA, drivers |
| [taeha-dispense-controller](taeha-dispense-controller/SKILL.md) | Taeha dispense SM, recipe, Modbus, safety |

## Career / Consulting

| Skill | What |
|-------|------|
| [30-resume-jd-workflow](30-resume-jd-workflow/SKILL.md) | JD→resume pipeline (migrated from rule) |
| [resume-jd-match](resume-jd-match/SKILL.md) | Match scoring and safe file outputs |
| [career-interview-prep](career-interview-prep/SKILL.md) | STAR / interview answers |
| [plm-systems-consulting](plm-systems-consulting/SKILL.md) | PLM / marine / defence consulting |

## Always-on Rules (not skills)

Remain in `.cursor/rules/` as `alwaysApply` or glob rules: `00-core-agent`, `10-model-orchestration`, `20-communication`, `40-technical-excellence`.

## Personal mirrors

Optional copies under `~/.cursor/skills/` for other repos. Prefer this project path when working in Codex hub.
