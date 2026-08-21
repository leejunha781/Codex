---
name: stm32-device-driver
description: Design, reconstruct, implement, and review STM32 BSPs, startup and boot code, bootloaders, boot-image pipelines, peripheral and external-device drivers, motor and sensor interfaces, GUI/input subsystems, secure update, and behaviorally equivalent firmware. Use for CMSIS reset/vector/linker work, ELF/HEX/BIN packaging, GPIO/EXTI, I2C/SPI/UART/RS-485/CAN, DMA/cache, timers/PWM/ADC, load cells, motor-driver ICs, BLDC/ACIM/stepper/servo drives, encoders, TouchGFX/LVGL, HAL/LL integration, firmware reconstruction, and material-dispensing hardware contracts.
---

# STM32 Device Driver

Build drivers as explicit contracts with deterministic ownership, timing, and failure behavior. Read `references/driver-contracts.md` before implementation.

- Read `references/firmware-reconstruction.md` for PDF schematic plus HEX/BIN reconstruction.
- Read `references/bsp-boot-images.md` for BSP, startup, linker/flash layout, bootloader, secure update, chain-loading, signing, and deterministic image generation.
- Read `references/input-subsystems.md` for GPIO/EXTI, buttons, key matrices, encoders, timer capture, ADC/DMA, wake, and safety inputs.
- Read `references/uart-servo-drive.md` for UART input to servo-drive command pipelines.
- Read `references/open-source-catalog.md` before adopting GitHub or other open-source material.
- Read `references/open-source-ecosystem-gaps.md` and `references/supply-chain-and-verification.md` for RTOS, security, middleware, test/debug, SBOM, vulnerability, fuzzing, and reproducible-build decisions.
- Read `references/bus-device-driver-patterns.md` for I2C, SPI, UART/RS-485, CAN/FDCAN, register codecs, DMA/cache, and concurrency.
- Read `references/motor-driver-ic-families.md` for brushed DC, BLDC/PMSM, ACIM, stepper, servo, and external motor-driver integration.
- Read `references/lcd-gui.md` for TouchGFX/LVGL display, touch, framebuffer, and single-owner GUI architecture.
- Read `references/adc-motor-calibration.md` for ADC/reference/current/voltage/back-EMF/temperature sensing and PWM-synchronous calibration.
- Read `references/drv8840.md` for TI DRV8840 PHASE/ENBL control, current limiting, faults, and the PROC-D100 IO board evidence.
- Invoke `$stm32-dispense-mass-control` for material-dependent dispensing, gravimetric feedback, shot-mass correction, density/temperature compensation, two-component ratio control, or calibration persistence. Invoke `$taeha-dispense-control-simulation` for plant models, toolchain selection, controller comparison, uncertainty sweeps, or HIL design. This driver skill owns hardware contracts; the dispense skill owns bounded control policy; the simulation skill owns executable models and evidence boundaries.
- Run `scripts/inspect_reconstruction_inputs.py --root <artifact-folder>` before analyzing supplied reconstruction artifacts.

## Required design record

- Hardware dependency and electrical assumptions
- Public API, units, ranges, return values, and timeout semantics
- Initialization and state transitions
- ISR/DMA/task ownership and synchronization
- Buffer ownership, overflow, and backpressure policy
- Error detection, recovery, safe state, and diagnostics
- Host-test seams and target-test procedure

## Implementation rules

- Keep register/HAL access in an adapter; keep parsers and state machines portable.
- Use fixed-width integer types and explicit unit suffixes such as `_ms`, `_us`, `_rpm`, `_ma`.
- For RS-485, define DE timing, turnaround, collision policy, frame timeout, CRC, and arbitration.
- Add one nearby `//YYYYMMDD Concise English change description` marker for each created or modified C/C++ logical block, using the actual date. Preserve earlier markers; in CubeMX files place markers only inside preserved user regions.
- For motors, default to disabled outputs, bounded commands, stale-command timeout, fault latching, and controlled recovery.
- For DMA/ring buffers, prove head/tail ownership and full/empty handling.
- Never invent pin mappings, timer channels, clock values, or polarity.
- For register devices, generate masks and codecs from the exact datasheet revision, serialize byte order explicitly, verify identity/configuration, and preserve reserved bits.
- For third-party code, pin tag and commit, verify file-level licenses/notices and transitive dependencies, record configuration/SBOM/local patches, and define upgrade and rollback ownership.
- For BSP/boot, prove vector/linker/flash layout, reset-safe outputs, image metadata, key custody, anti-rollback, power-loss recovery, and deterministic manifests.
- For GUI, keep one GUI owner and isolate control/safety; prove framebuffer memory, cache/DMA, transfer completion, and touch mapping.
- For ADC feedback, separate internal/reference/channel/fixture/runtime calibration and reject stale, implausible, or invalid coefficients.
- For input subsystems, define electrical state, polarity, filtering/debounce, ownership, overflow/freshness, low-power wake, and fail-safe fault behavior.

## Reconstruction boundary

- Produce a clean, reviewable, behaviorally equivalent implementation. Never promise recovery of original names, comments, types, macros, file boundaries, compiler settings, or undefined behavior from HEX.
- Classify every conclusion as `CONFIRMED`, `INFERRED`, or `UNKNOWN`, with its source artifact or measurement.
- Require the servo-drive/controller manual and wiring interface when the motor sheet does not define command signaling. Do not infer pulse/direction, PWM, analog, UART, RS-485, CANopen, EtherCAT, enable, brake, or alarm semantics from motor ratings.
- Keep drive outputs disabled and bench power current-limited until polarity, electrical levels, limits, fault inputs, and stop behavior are confirmed.

## Handoff

Use `$embedded-code-audit` for concurrency, bounds, integer math, calibration/NVM integrity, and safe-state review. Use `$embedded-test-debug` for host, target, HIL, and physical-process tests.

For substantive implementation or reconstruction, follow `$stm32-collaborative-development` and its review protocol. Obtain explicit approval before transmitting a minimal frozen packet to Cursor or Claude and report actual participation.

## Project Memory traceability

Whenever this skill changes code, update the registered PMC vault in the same task with repository-relative paths, intent, affected contract, verification, assumptions, residual risks, branch/revision, and reviewer participation. Never store source dumps, absolute paths, secrets, or transcripts; route semantic conflicts through the Promotion Inbox.
