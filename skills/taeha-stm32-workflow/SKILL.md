---
name: taeha-stm32-workflow
description: Run end-to-end Taeha STM32 controller development and release delivery. Use for BSP/startup/linker and secure boot, deterministic image generation, GPIO/EXTI/input capture, PDF/HEX reconstruction, IAR EWARM and CubeMX/CubeIDE, UART/RS-485/CAN, motor-driver ICs and servo/BLDC/stepper control, ADC calibration, TouchGFX/LVGL, material dispensing, board bring-up, requirements traceability, open-source qualification, independent review, and release-quality evidence.
---

# Taeha STM32 Workflow

Use a requirements-first, hardware-aware loop. Preserve CubeMX-generated regions and never claim hardware behavior without board evidence.

This skill extends `$stm32-collaborative-development` with Taeha-specific engineering and release rules. When loaded by the STM32 skill, merge its requirements, audit, test, driver, dispensing, and simulation evidence into one traceability contract; do not treat this as a separate generic workflow.

## Boot

1. Read `references/taeha-jd-profile.md`.
2. Identify MCU/board, clock tree, HAL/LL version, RTOS, toolchain, pin map, peripherals, external circuits, protocol documents, and acceptance tests.
3. If the exact board is unknown, keep hardware access behind interfaces and state assumptions explicitly.
4. Run `scripts/inspect_stm32_project.ps1 -ProjectRoot <path>` before editing an existing project.
5. For reconstruction, read `references/reconstruction-delivery.md`, then invoke `$stm32-device-driver` and its reconstruction references.
6. When reconstruction artifacts are outside the firmware project, run the device-driver artifact inspector on that external folder separately; project-root discovery does not imply sibling-folder discovery.
7. Invoke `$stm32-device-driver` and read only the relevant BSP/boot, input, bus, motor-IC, GUI, ADC, reconstruction, or supply-chain reference before designing that subsystem.
8. For motor-driven dispensing, load-cell integration, material-mass calibration, shot correction, density/temperature compensation, or two-component mixing, invoke `$stm32-dispense-mass-control` before defining the control policy.
9. Invoke `$taeha-dispense-control-simulation` before selecting MATLAB/Simulink or an alternative, fitting a plant, tuning delayed feedback, comparing controllers, or claiming simulation/SIL/PIL/HIL evidence.
10. Load `$embedded-code-audit` for review gates and `$embedded-test-debug` for the verification ladder; for substantive or safety-relevant work, keep `$stm32-collaborative-development` as the orchestration layer.

## Engineering loop

1. Convert the request into testable requirements and a traceability table.
2. Separate portable domain logic from STM32 HAL/LL adapters.
3. Define timing, concurrency, ISR/task ownership, buffer lifetimes, fault states, and recovery before implementation.
4. Implement the smallest vertical slice. Keep CubeMX user code inside `USER CODE BEGIN/END` regions or separate files. Mark each created or modified C/C++ logical block with `//YYYYMMDD Concise English change description` using the actual date.
5. Build host tests for pure logic. Build the cross target when the Arm toolchain and generated project are available.
6. Run compiler warnings, clang-tidy/cppcheck when installed, and `$embedded-code-audit`.
7. When third-party code is present, verify pinned provenance, licenses/notices, transitive dependencies, component inventory/SBOM, security status, reproducible configuration, local patches, and upgrade/rollback ownership.
8. Verify on hardware with UART logs, logic-analyzer captures, current limits, fault injection, and reset/recovery tests.
9. Report changed files, tests, dependency provenance, unverified hardware assumptions, and recovery steps.

## Extended subsystem gates

- **I2C/SPI/CAN devices:** approve a register/transport contract, source revision, license, bus ownership, timeout/recovery, and logic-analyzer test before integration.
- **Motor-driver ICs:** prove reset-safe outputs, exact suffix/register map, PWM mode/polarity, current sensing, fault decode, and recovery with the power stage initially disabled.
- **LCD GUI:** prove board/display/touch/memory/cache bring-up before framework integration; use one GUI owner and bounded messages from confirmed controller state.
- **ADC compensation:** prove ADC/reference/channel calibration, PWM-trigger timing, coefficient validity, sample freshness, sensor plausibility, and independent protection before closed-loop enable.
- **BSP/boot images:** prove reset-safe board state, startup/vector/linker/flash map, image descriptor/signing policy, key custody, power-loss recovery, rollback, and deterministic manifests.
- **Inputs:** prove electrical levels/pulls/polarity, debounce/filtering, ISR ownership, overflow/freshness, wake, and stuck/open/short fault behavior.
- **Open-source middleware:** pin tag and commit, verify file-level licenses and transitive components, retain notices, record SBOM/configuration/local patches, and assign upgrade/rollback ownership.
- **USB/network/storage/serialization:** prove bounded buffers, allocation failure, ISR/task handoff, cache/DMA coherency, malformed-input handling, backpressure, persistence compatibility, and host fuzz seams.
- **Virtual targets:** use Renode or another model only for implemented behavior; timing, analog, DMA/cache, electrical, and safety claims remain `UNVERIFIED` until measured.

## Multi-engine review

For substantive or safety-relevant work, follow `stm32-collaborative-development/references/multi-engine-review.md`. Obtain explicit task-specific approval before transmitting the frozen packet; keep first passes independent and report actual participation.

## Reconstruction delivery gate

1. Inventory and hash supplied artifacts; confirm analysis authorization and exact MCU/board revision.
2. Build a `CONFIRMED`/`INFERRED`/`UNKNOWN` evidence matrix for clocks, pins, peripherals, protocols, motor interface, units, timing, calibration, and faults.
3. Define the required equivalence level before implementation: build, peripheral, protocol, motor sequence, or fault parity.
4. Stop hardware-driving work when enable/brake/direction polarity, electrical levels, limits, or safe state remain unknown. Continue only with interface-isolated code and host tests.
5. Generate an IAR linker map and preserve firmware hashes, compiler version, project options, board revision, captures, and test evidence.

## Taeha priorities

- Prioritize deterministic state machines, serial robustness, motor interlocks, watchdog strategy, diagnostics, and serviceability.
- Treat dispenser and stepper-domain behavior as a validated domain only when `$stm32-dispense-mass-control` has produced a profile-specific, bounded control contract and physical evidence. Otherwise mark accuracy, repeatability, ratio, timing, and safe-state claims `UNVERIFIED`.
- Use STM32F303 only as an initial compatibility baseline; do not assume it is the Taeha production MCU.
- Reject blocking waits in ISRs and unbounded waits in control paths.

## Exit criteria

- Requirements link to code and tests.
- Host tests pass; target build status is explicit.
- No unresolved critical audit finding.
- Every adopted third-party component has pinned provenance, license/notice evidence, reproducible configuration, and upgrade/rollback ownership.
- Every created or modified C/C++ logical block has the required dated English marker in a regeneration-safe location.
- Any code change has the required same-task PMC traceability update, or the unavailable vault is reported explicitly.
- Hardware-only claims are labeled `UNVERIFIED` until measured.
- CubeMX regeneration risk is documented.
