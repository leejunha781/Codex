---
name: embedded-test-debug
description: Plan and execute STM32 firmware verification and debugging. Use for BSP/startup/linker and bootloader/image validation, input subsystems, reconstructed-firmware equivalence, UART/RS-485/CAN and DMA gateways, motor-driver ICs, ADC calibration, TouchGFX/LVGL, material dispensing and shot-mass correction, load-cell/calibration persistence, host unit/sanitizer/fuzz tests, virtual targets, cross builds, ST-LINK/pyOCD/OpenOCD, instrumented measurements, HIL, fault injection, and release qualification.
---

# Embedded Test and Debug

This is the single canonical `$embedded-test-debug` skill. When invoked through `$stm32-collaborative-development`, extend the STM32 task with the verification and debugging rules below. Load `$taeha-stm32-workflow` for Taeha end-to-end delivery context and `$taeha-dispense-control-simulation` for model, SIL, PIL, HIL, or physical-evidence planning. Feed failures back through `$embedded-code-audit` and keep one shared evidence ladder.

Read `references/test-debug-matrix.md`, then climb the test ladder without skipping evidence.

For STM32 dispensing, invoke `$stm32-dispense-mass-control`, read `references/stm32-dispense-test-debug.md`, and use `references/open-source-test-tooling.md` to select reproducible tools. Trace peripheral behavior through `$stm32-device-driver` and audit results through `$embedded-code-audit`.

## Test ladder

1. Static compile with maximum practical warnings plus version-pinned static analysis.
2. Host tests for parsers, codecs, state machines, limits, timeout wraparound, and faults.
3. Run sanitizers and coverage-guided fuzzing for portable code when supported; preserve minimal failing inputs.
4. Cross build and link-map review for flash/RAM/stack budgets and deterministic artifacts.
5. Run virtual-target tests only for implemented CPU/peripheral behavior and record unsupported claims.
6. On-target smoke test with outputs initially disabled.
7. Peripheral tests with oscilloscope or logic-analyzer evidence.
8. HIL/fault injection for cable loss, CRC/malformed frames, queue exhaustion, sensor range, driver fault, brownout/storage loss, and watchdog reset.
9. Regression report with firmware hash, component inventory/SBOM, board revision, tool versions, corpus, and artifacts.

For material-dispensing simulation, invoke `$taeha-dispense-control-simulation`; preserve deterministic seeds, profiles, raw CSV/JSON, plots, model revisions, uncertainty sweeps, and separate `SIMULATED`, `SIL`, `PIL`, `HIL`, and `PHYSICAL` claims.

For DRV8840 targets, test outputs-disabled reset, at least 1 ms wake settling, both directions, ENBL PWM, slow/fast decay stop behavior, nFAULT assertion, OCP reset recovery, current-limit accuracy, encoder polarity, and ADC current feedback before closed-loop operation.

For register devices, capture identity/configuration reads, bus mode/frequency/address or chip select, byte order, timeout, retry, recovery, wrong-ID, and corrupted/truncated replies.

For LCD GUI, test boot-to-first-frame, color/touch grids, orientation, flush/vsync timing, DMA/cache, queue pressure, memory high-water marks, corrupted assets, missing devices, soak, and motor-fault presentation without safety coupling.

For ADC motor compensation, test self-calibration, VREFINT/VDDA, zero/gain fixtures, coefficient CRC/version, PWM-synchronous timing, independent instrument correlation, drift/temperature, saturation, open/short, DMA freeze, and save brownout.

For BSP/boot images, test vector/linker placement, reset-safe outputs, erased/truncated/oversized/wrong-target/bad-signature/rollback images, readback, watchdog/reset causes, handoff, and reset injection at every flash transition.

For inputs, test floating/open/short/stuck, bounce/tick wrap, min pulse/max rate, simultaneous events, overflow, EXTI storms, encoder illegal transitions, capture overrun/no-signal, DMA freeze, wake, and reset-time interlocks.

For third-party middleware, test the pinned configuration, allocation failure, malformed/oversized/duplicate inputs, restart/migration, transitive versions, patches, and license/notice inventory. Mocks prove only explicit host seams.

For Renode, pyOCD, or OpenOCD, pin tool and target configuration, record probe/transport settings, use non-interactive scripts, and repeat a physical-board subset before release.

## Debug discipline

- Reproduce before changing code; preserve logs and minimal stimulus.
- Separate compile, link, flash, connection, clock, protocol, and application failures.
- Check reset reason, fault registers, stack watermark, interrupt load, and jitter.
- Do not disable watchdogs or safety interlocks except in a controlled test with explicit restoration.
- Never flash a connected production board without confirmation of target, probe serial, image, and recovery method.

## Exit criteria

- Every requirement has a test or documented justification.
- Host tests pass and target status is explicit.
- No unexplained reset or latched fault remains.
- Measurements identify board revision and firmware build.
- Dispense mass, repeatability, A/B ratio, safe stop, and fault recovery remain `UNVERIFIED` until measured on the exact apparatus with the profile and calibration revision recorded.
- Any C/C++ code created or corrected during debugging has one actual-date change marker per logical change in a regeneration-safe location.

For substantive qualification, follow `$stm32-collaborative-development` and its review protocol. Obtain explicit approval before transmitting the frozen packet, keep first passes independent, and report actual participation.

## Project Memory traceability

Whenever testing or debugging changes code, update the registered PMC vault with repository-relative paths, objective/fault, rationale, evidence, assumptions, residual risks, branch/revision, and reviewer participation. Do not store full source, absolute paths, credentials, secrets, or transcripts.
