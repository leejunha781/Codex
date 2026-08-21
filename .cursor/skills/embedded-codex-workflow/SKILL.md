---
name: embedded-codex-workflow
description: Orchestrates Cursor + Claude + Codex for STM32/embedded and Taeha dispense firmware. Use when starting embedded work with Codex, multi-model firmware pipelines, driver PRs, or linking Cursor skills to Codex implementation.
---

# Embedded ↔ Codex Development Workflow

## Goal

Ship embedded firmware (STM32 drivers, Taeha dispense controller logic) with the same discipline as the career hub: structure → implement → verify → close.

For whole-workspace Max Mode routing, also load `cursor-pro-max-orchestration`.

## Model Split

| Layer | Tool | Role |
|-------|------|------|
| Spec / architecture / RCA | Claude Opus/Sonnet (**Max Mode**) | State machines, timing budgets, trade-offs |
| Implementation | GPT-5.3 Codex / Cursor Agent | Drivers, HAL wrappers, tests, refactors |
| Second-pass review | Claude or Bugbot | Safety, ISR budget, interface correctness |
| Repo / skills / PR | Cursor Agent | `.cursor/skills`, git, Cloud Agents |

Do not pay twice for the same subtask unless a second pass improves quality. Prefer **Max Mode** for firmware design and RCA.

## Skill Loading Order

1. `cursor-pro-max-orchestration` (session routing)
2. This skill (`embedded-codex-workflow`)
3. Domain: `taeha-dispense-controller` and/or `stm32-device-drivers`
4. Project rules already always-on (core, communication, models)

## Session Bootstrap Checklist

```
Task Progress:
- [ ] 1. Restate goal, MCU part, constraints, success criteria
- [ ] 2. Locate Datasheet / RM / Errata / schematics / I/O map
- [ ] 3. Load domain skill(s); note accuracy boundaries
- [ ] 4. Agree deliverables (files, tests, measurements)
- [ ] 5. Implement smallest vertical slice
- [ ] 6. Bench/HIL evidence before claiming done
- [ ] 7. Commit focused change; update MEMORY if durable fact learned
```

## Codex Prompt Template

Paste into Codex / Cursor Agent when starting a firmware task:

```markdown
## Goal
[One sentence]

## MCU / Board
- Part: [e.g. STM32Fxxx]
- Clock: [HSE/HSI + PLL]
- Toolchain: [CubeIDE / arm-none-eabi / existing]

## Constraints
- Match existing driver style in [path]
- ISR budget: [µs]; control period: [ms]
- No new RTOS unless asked

## Skills to follow
- @.cursor/skills/cursor-pro-max-orchestration/SKILL.md
- @.cursor/skills/embedded-codex-workflow/SKILL.md
- @.cursor/skills/stm32-device-drivers/SKILL.md
- @.cursor/skills/taeha-dispense-controller/SKILL.md (if dispense)
- @MEMORY.md

## Deliverables
- [ ] Driver .h/.c or SM module
- [ ] Integration notes (IRQ/DMA/CubeMX)
- [ ] Test steps + expected scope/log evidence

## Out of scope
[List]
```

## Vertical Slice Order

Prefer end-to-end thin path over wide stubs:

1. Heartbeat GPIO / UART log
2. One actuator or sensor path
3. One host command → one shot / motion
4. Fault path + safe-state
5. Recipe/calibration polish

## Verification Bar

- Build succeeds in project toolchain
- At least one measured check (scope toggle, log timestamp, unit test)
- Fault injection for the changed path when safety-relevant
- No invented Taeha-internal architecture

## Durable Learning Loop

When a fact will matter again (part number, pin map convention, protocol quirk):

1. Update `MEMORY.md` briefly
2. Or extend the domain skill `reference.md` (prefer skill for technical patterns)
3. Keep accuracy boundaries honest (LPC1769 vs STM32F)

## Related

- [examples.md](examples.md) — sample task briefs
- `stm32-device-drivers`, `taeha-dispense-controller`
- Notion hub: Taeha controller interview/prep page (user)
