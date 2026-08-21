---
name: stm32-device-drivers
description: STM32 firmware and device-driver development with CubeMX/HAL/LL, register-level peripherals, ISR/DMA, motors, and serial I/O. Use when writing or reviewing STM32 code, drivers, HAL wrappers, CubeMX projects, or Cortex-M bring-up.
---

# STM32 & Device Driver Development

## When to Use

- New STM32 peripheral driver or HAL/LL wrapper
- CubeMX / STM32CubeIDE project setup or regeneration
- ISR, DMA, timer, ADC, UART/SPI/I2C, GPIO bring-up
- Motor (DC/BLDC/STEP/SERVO) low-level control paths
- HardFault / timing / jitter debug on Cortex-M

Load [reference.md](reference.md) for register/ISR patterns and checklists. For Taeha dispense product context, also load `taeha-dispense-controller`. For Codex orchestration, load `embedded-codex-workflow`.

## Accuracy Boundaries (User Context)

From MEMORY / Notion hub (corrected):

| Claim | Fact |
|-------|------|
| KSS-III MCU | **NXP LPC1769** (Cortex-M3) — not STM32 |
| Intellian | **STM32F** existing BLDC code analysis / validation / tuning |
| Transferable | Exception/ISR model, Embedded C, state machines, async comms, HW/FW debug |
| Must re-verify per part | Clock tree, pinmux, DMA, startup/linker, vendor libs, errata |

Never invent part numbers. Ask for or read Datasheet + Reference Manual + Errata for the target STM32.

## Default Stack

1. **Preferred:** STM32CubeMX → HAL (or LL where latency/size matters) → thin driver layer → app state machine
2. **Bare-metal OK** when repo already uses register drivers — match existing style
3. **RTOS** only if project already uses FreeRTOS/ThreadX; do not introduce without ask

## Driver Layering

```
App / State machine
  → Device service (dispense, motor, protocol)
    → Driver API (init, start, stop, irq_handler hooks)
      → HAL/LL or register access
        → Hardware
```

Rules:

- Keep HAL calls inside the driver; app must not sprinkle `HAL_*` everywhere
- One peripheral ownership; document who owns DMA/IRQ
- `volatile` for MMIO/shared flags — does **not** imply atomicity or barriers
- ISR: clear flag → timestamp/capture → enqueue; no `printf`, `malloc`, blocking, long loops

## Implementation Workflow

```
Task Progress:
- [ ] 1. Identify MCU part, package, clock, pin map
- [ ] 2. Read RM sections for target peripheral + DMA + NVIC
- [ ] 3. Define public driver API (init / deinit / start / stop / ioctl)
- [ ] 4. Implement + wire IRQ/DMA callbacks
- [ ] 5. Unit/bench smoke: clock, GPIO loopback, peripheral self-test
- [ ] 6. Measure timing budget (ISR entry → done, DMA half/full)
- [ ] 7. Document assumptions, errata workarounds, test evidence
```

## Code Conventions

- C11; headers self-contained; includes at top of file
- Naming: `drv_<periph>_init`, `drv_<periph>_irq_handler`
- Return `int`/`enum` error codes; assert only in debug builds
- Separate `*_config.h` for pins/clocks/timeouts — no magic numbers in `.c`
- Prefer explicit units in names: `_us`, `_hz`, `_mv`

## Debug Order

1. Power → Clock → Reset → Pinmux → Peripheral enable
2. Scope/logic on critical pins before blaming firmware
3. HardFault: CFSR/HFSR/MMFAR/BFAR + stacked PC/LR/xPSR + SP watermark
4. DMA: buffer ownership, cache (if applicable), circular vs normal, overrun

## Output Expectations

- Compilable driver pair (`.h`/`.c`) matching project tree
- Short integration notes: CubeMX settings, IRQ priorities, DMA channels
- Test steps with expected measurements (scope/log)
- Korean chat ↔ Korean; code comments in English unless repo is Korean

## Related

- [reference.md](reference.md) — patterns, ISR/DMA budget, motor notes
- Skill `taeha-dispense-controller` — product state machine / I/O
- Skill `embedded-codex-workflow` — Cursor ↔ Codex pipeline
