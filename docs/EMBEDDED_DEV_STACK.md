# Embedded Dev Stack — Cursor Pro + Claude + Codex

Playbook for STM32 / Taeha dispense firmware using this hub.

## 1. One-time Setup (User)

1. Cursor Pro: enable Claude Opus/Sonnet, GPT-5.3 Codex, **Max Mode**  
2. Claude Pro Project **"Joonha Career & Codex"**: upload `MEMORY.md` + embedded skill `SKILL.md` files  
3. Open this repo; confirm `.cursor/skills/` listed under Settings → Rules/Skills  
4. Cloud Agents environment linked to `leejunha781/Codex`

Details: [SETUP_CURSOR_PRO_CLAUDE_CODEX.md](SETUP_CURSOR_PRO_CLAUDE_CODEX.md)

## 2. Every Firmware Session

```
@MEMORY.md
@.cursor/skills/cursor-pro-max-orchestration/SKILL.md
@.cursor/skills/embedded-codex-workflow/SKILL.md
@.cursor/skills/stm32-device-drivers/SKILL.md
@.cursor/skills/taeha-dispense-controller/SKILL.md   # if dispense
```

Then state: MCU part, board, goal, constraints, success criteria.

## 3. Model Hand-offs

| Step | Who | Output |
|------|-----|--------|
| Design SM / timing / safety | Claude Max | Mermaid or checklist |
| Implement drivers / app | Codex in Cursor | `.c`/`.h`, Cube notes |
| Measure / RCA | Cursor + scope/log | Evidence |
| Polish design note | Claude | Short review |

## 4. Accuracy Boundaries

- KSS-III product MCU: **NXP LPC1769** (not STM32)  
- Intellian: STM32F **existing** BLDC code validation/tuning  
- Do not invent Taeha private schematics/firmware  

## 5. Learning Loop

New pin map / protocol quirk / calibration rule → update domain skill `reference.md` and/or `MEMORY.md`.
