---
name: embedded-code-audit
description: Perform rigorous embedded C/C++ review for STM32 firmware. Use for BSP/startup/linker and bootloaders, deterministic boot images, input subsystems, reconstructed firmware, UART/RS-485/CAN and DMA gateways, motor-driver ICs, ADC calibration, TouchGFX/LVGL, material dispensing and shot-mass correction, load-cell/calibration persistence, static/formal analysis, ISR/RTOS concurrency, memory safety, protocol parsers, middleware, supply-chain evidence, motor safety, and release gates.
---

# Embedded Code Audit

This is the single canonical `$embedded-code-audit` skill. When invoked through `$stm32-collaborative-development`, extend the STM32 task with the audit rules below. Load `$taeha-stm32-workflow` for Taeha end-to-end delivery context and `$taeha-dispense-control-simulation` when executable models, generated control, or simulation evidence are in scope. Reconcile their requirements into the same finding and evidence model rather than duplicating skill definitions.

## Review order

1. Establish MCU, compiler, optimization, RTOS, interrupt priorities, and hardware safety context.
2. Read diffs plus surrounding call paths.
3. Apply `references/code-audit-checklist.md`.
4. For STM32 dispensing, read `references/stm32-dispense-audit.md`, invoke `$stm32-dispense-mass-control`, and trace hardware contracts through `$stm32-device-driver`.
5. Read `references/open-source-audit-tooling.md` before selecting static or formal tools.
6. Run compiler warnings, the project-selected static/formal tools, and tests when available.
7. Report findings by severity with exact file/line, failure scenario, and smallest safe fix.

## Severity

- P0: can energize unsafe hardware, corrupt firmware, or defeat critical protection.
- P1: likely crash, deadlock, memory corruption, protocol loss, or uncontrolled motor behavior.
- P2: correctness or maintainability defect with bounded impact.
- P3: improvement with no demonstrated defect.

## Mandatory checks

- Bounds, conversions, shifts, wraparound, division, and scaling.
- ISR/task races, priority inversion, atomicity, and callback lifetime.
- Blocking ISR/control-loop calls, unbounded retries, and timeout wraparound.
- Parser validation and CRC before command execution.
- Watchdog feed location and fault escalation.
- Safe startup, brownout, comms loss, sensor fault, and driver-fault behavior.
- For DRV8840, verify ENBL/PHASE semantics by physical pin number, open-drain nFAULT biasing, nRESET/nSLEEP startup state, wake delay, OCP latch recovery, DECAY brake/coast behavior, and bounded VREF/current-code calculations.
- HAL return values, partial initialization cleanup, and CubeMX regeneration safety.
- For dispensing, verify profile isolation, sensor-quality gates, bounded corrections, unit consistency, fixed-point saturation, transport-delay assumptions, A/B ratio observability, flash endurance, version/CRC/atomicity, and fail-safe rollback. A learned value may never override a hard current, pressure, temperature, motion, ratio, timeout, or interlock limit.
- For simulation or generated control, invoke `$taeha-dispense-control-simulation` and verify sample times, delay units, saturation, anti-windup, fixed-point ranges, solver failure/fallback, model-code equivalence, deterministic artifacts, and honest `SIMULATED`/`SIL`/`PIL`/`HIL`/`PHYSICAL` labels.
- I2C/SPI/CAN register encoding, byte order, reserved-bit policy, bus ownership, retries, timeout distinction, DMA/cache coherency, and recovery.
- Motor-driver exact-suffix reset values, PWM/deadtime/break ownership, current sensing, configuration readback, first-fault preservation, bounded clear, and no automatic re-enable.
- ADC trigger/sample timing, reference/offset/gain provenance, checked scaling, coefficient version/CRC, freshness, saturation/plausibility, and calibration-state power limits.
- GUI single-owner enforcement, callback bounds, queue overflow, framebuffer/cache/DMA/flush correctness, memory bounds, and separation from control/safety.
- BSP/startup vector order, linker regions, runtime initialization, reset-safe pins, clock/watchdog/reset cause, TrustZone/core/cache, and regeneration boundaries.
- Boot-image range arithmetic, flash geometry/ECC, authenticated metadata, key separation, anti-rollback, power-loss-safe transition, revert, jump preconditions, deterministic artifacts, and readback.
- Input pulls/polarity, ISR clearing/deferral, debounce wraparound, key ghosting, encoder/capture rollover, overrun/freshness, wake, and fail-safe interlocks.
- Third-party pinned provenance, file-level licenses/notices, transitive dependencies, security advisories, generated-code provenance, local patches, SBOM, reproducible configuration, and upgrade ownership.
- Middleware bounds and ownership for USB, network buffers, filesystems, serialization, CAN transfer IDs, malformed/truncated/duplicate/replayed frames, allocation failure, and backpressure.
- Verification quality: real host seams, sanitizer/fuzz coverage, preserved failing corpora, and no analog/timing/cache/electrical claims from mocks or virtual targets.
- Change traceability: each created or modified C/C++ logical block has the actual-date marker in a regeneration-safe location.

If no actionable defect exists, say so and list residual hardware-validation risks.

For substantive or safety-relevant audits, follow `$stm32-collaborative-development` and its review protocol. Obtain explicit approval before transmitting the frozen packet; keep Cursor and Claude first passes independent and reconcile them against primary evidence.

## Project Memory traceability

If an audit causes code changes, update the registered PMC vault in the same task with repository-relative paths, finding-to-fix rationale, verification, assumptions, residual risks, branch/revision, and actual reviewer participation. Never store source dumps, absolute paths, secrets, or transcripts; route conflicts through the Promotion Inbox.
