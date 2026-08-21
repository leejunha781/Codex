---
name: stm32-dispense-mass-control
description: Design, implement, audit, calibrate, and test safe STM32 dispensing control with motor actuation, load-cell feedback, material-, temperature-, pressure-, nozzle-, and lot-dependent compensation, shot-to-shot learning, two-component ratio supervision, and calibration persistence. Use for gravimetric or volumetric pumps, augers, syringes, valves, steppers, servos, BLDC drives, and adhesive, resin, solder-paste, or chemical dispensing.
---

# STM32 Dispense Mass Control

Build material-aware dispensing as a bounded safety-critical control system. This skill restores the missing bridge between STM32 motor/device drivers and measured shot-mass accuracy. Read `references/control-architecture.md` and `references/calibration-and-persistence.md` before implementation. Read `references/evidence-catalog.md` before adopting code or control ideas. Invoke `$taeha-dispense-control-simulation` for plant identification, Python/OpenModelica/Renode/Simulink selection, controller comparison, Monte Carlo, SIL/PIL/HIL design, or learning materials. Use `$embedded-code-audit` and `$embedded-test-debug` before release.

## Evidence gate

1. Confirm the exact MCU, board revision, schematic, motor/drive interface, valve behavior, load-cell/ADC chain, temperature and pressure sensors, nozzle/mixer, material/component/lot, toolchain, HAL/LL/RTOS versions, and acceptance limits.
2. Record every unknown pin, polarity, limit, calibration constant, density curve, viscosity range, safe state, and timing as `UNKNOWN`. Keep outputs disabled until electrical and mechanical safe states are confirmed.
3. Separate claims as `CONFIRMED`, `INFERRED`, `UNKNOWN`, or `UNVERIFIED`. Simulation, host tests, and community projects do not prove physical mass accuracy.
4. Define units and reference conditions for every quantity. Raw ADC counts are not mass; motor steps are not volume; total mass does not prove a two-component ratio.

## Required architecture

Keep portable domain logic independent from STM32 adapters:

- `material_profile`: material/component/lot, density-vs-temperature, viscosity or process class, pot life, nozzle/mixer, pump/gear ratio, calibration revision, and valid operating envelope.
- `shot_planner`: target mass to bounded precharge, motion profile, cutoff, valve timing, decompression or suck-back, settle, purge, and recovery commands.
- `mass_estimator`: tare, calibration, filtering, stability, saturation, stale-sample, drift, creep, vibration, outlier, and confidence state.
- `bounded_compensator`: accepted-shot-only correction with gain, rate, absolute, sample-count, and profile-match limits.
- `ratio_supervisor`: independently measured A/B mass or qualified per-component flow sensors, startup lag, purge, ratio window, pot life, and mixer-pressure checks. Motor step/time proxies are `INFERRED` process indicators only and cannot pass or certify ratio unless a profile-specific validation establishes the complete measurement chain.
- `calibration_store`: versioned schema, units, profile key, provenance, CRC, sequence number, transactional or dual-copy update, and factory fallback.
- `safety_state_machine`: disabled, homing/prime, ready, dispensing, settling, complete, recoverable fault, and latched fault states.
- STM32 adapters: timer/PWM/stepper/servo/RS-485 motor output; ADC/SPI/DFSDM load-cell input; temperature/pressure input; NVM; monotonic time; diagnostics.

## Control hierarchy

Apply these layers in order; a later layer may never bypass an earlier one.

1. **Hard envelope:** current, torque, speed, acceleration, travel, pressure, temperature, ratio, timeout, stale command, interlock, and safe-output limits.
2. **Feed-forward baseline:** profile-specific map from target mass and measured conditions to motor/valve command, cutoff, precharge, decompression, and settle time.
3. **In-shot feedback:** enable only when sensor bandwidth, transport delay, mechanical response, and sample quality are measured to support it.
4. **Shot-to-shot correction:** prefer this for delayed viscous or thixotropic flow. Update only after a stable, valid, profile-matched measurement.
5. **Optional learning:** bounded EWMA, integral, iterative-learning, or identified-model updates require explicit confidence, rollback, and disable criteria.

Example policy is pseudocode, not a universal tuning rule. Require `target_mass_mg >= validated_min_target_mg`; implement the equivalent with explicit fixed-point scale, widened intermediates, saturating operations, and divide guards when floating point is not justified:

```text
e_k = target_mass_mg - accepted_mass_mg
gain_next = clip(gain_k * (1 + alpha * e_k / target_mass_mg), gain_min, gain_max)
gain_next = rate_limit(gain_next, gain_k, max_delta_per_shot)
```

Reject the update on sensor invalid/stale/saturated, unstable scale, outlier, timeout, profile mismatch, purge/prime shot, manual intervention, fault, target below validated minimum, or insufficient samples. Reset or quarantine learning after material, lot, nozzle, mixer, pump, gear ratio, density curve, temperature band, or firmware/calibration-schema change.

## Load-cell and process rules

- Require tare plus known-mass two-point calibration at minimum; use multi-point calibration when nonlinearity is measured.
- Validate excitation/reference, ADC data rate and settling, vibration, creep, zero drift, temperature drift, mechanical preload, tilt, container contact, saturation, and recovery.
- Model residual dribble, transport delay, pressure settling, valve lag, backlash, compressibility, air bubbles, cavitation, empty supply, blockage, and motor slip or missed steps.
- For two-component material, supervise A and B individually when possible. Include density-versus-temperature, startup transient, purge volume, ratio tolerance, pot life, and mixer pressure.
- Never carry learned correction across incompatible profiles automatically.

## STM32 implementation rules

- Use `$stm32-device-driver` for exact timer, ADC, DMA, SPI, UART/RS-485, motor-drive, sensor, and NVM contracts.
- Use wrap-safe monotonic time, fixed-width types, explicit unit suffixes, saturating or proven arithmetic, deterministic ownership, and nonblocking ISR/control paths.
- Define cache/DMA coherency, ISR/task handoff, sample timestamps, buffer overflow policy, watchdog ownership, brownout behavior, and reset-safe outputs.
- Never erase or program flash in a dispense ISR or deadline-critical control task. Queue a bounded `persist_pending` request for a non-control context and define queue-full, persist-failure, reset, and mid-sequence behavior. A shot may complete safely even when learning remains dirty and unpersisted.
- On cache-less STM32 parts, still prove DMA buffer ownership and ordering. On F7/H7 or other D-Cache parts, place buffers in a coherent region or perform correctly aligned clean/invalidate operations with barriers; test the production linker, MPU, and cache configuration.
- Preserve CubeMX user regions. Add one nearby dated English marker to each modified C/C++ logical block: `//YYYYMMDD Concise English change description`.
- Do not modify vendor or open-source code merely to add a marker.

## Verification gate

Read `references/verification-matrix.md`; it is the canonical dispensing verification matrix. At minimum perform static analysis, host unit/property/fuzz tests, fixed-point and tick-wrap tests, captured-sensor replay, target build/link-map review, outputs-disabled target smoke test, ADC/DMA timing measurement, calibration-store corruption/brownout tests, HIL fault injection, and statistically meaningful repeated-shot tests. Report bias, spread, confidence, rejected-shot criteria, raw evidence, firmware hash, board revision, profile/calibration revision, and environmental conditions. Label mass accuracy, repeatability, ratio accuracy, safe stop, and recovery `UNVERIFIED` until measured on the exact apparatus.

## Exit criteria

- Requirements trace to code, calibration records, and tests.
- Compensation is bounded, profile-isolated, reversible, and subordinate to safety limits.
- Invalid measurements cannot update learned state or energize unsafe outputs.
- Calibration persistence survives reset and interrupted writes without silently accepting corrupt data.
- Open-source provenance, license compatibility, configuration, patches, SBOM impact, upgrade owner, and rollback are recorded.
- `$embedded-code-audit` has no unresolved critical finding; `$embedded-test-debug` reports host, target, and HIL evidence separately.
