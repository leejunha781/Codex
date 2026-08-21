---
name: taeha-dispense-control-simulation
description: Model, simulate, identify, tune, and verify material-dispensing control for Taeha STM32 systems without assuming MATLAB. Use for material/nozzle/temperature feed-forward, sensor-delay and dribble models, servo/stepper/DC/BLDC/AC actuator abstractions, load-cell feedback, bounded shot-to-shot or run-to-run correction, A/B ratio supervision, profile isolation, CRC/dual-copy calibration, Python/OpenModelica/Renode/Simulink selection, SIL/PIL/HIL planning, Monte Carlo tests, and control-theory learning.
---

# Taeha Dispense Control Simulation

Treat simulation as executable evidence with a declared fidelity boundary. Never convert a simulated result into a physical accuracy or safety claim.

This skill extends `$stm32-collaborative-development` with Taeha dispensing-model and control-learning rules. When loaded by the STM32 skill, contribute model assumptions, controller evidence, and fidelity labels to the shared task contract; do not create an independent STM32 evidence path.

## Boot

1. Invoke `$stm32-dispense-mass-control` for the control policy, `$stm32-device-driver` for exact motor/sensor/NVM contracts, `$embedded-code-audit` for firmware/model-code review, and `$embedded-test-debug` for target and HIL evidence. Use `$stm32-collaborative-development` as the parent orchestration skill for substantive STM32 work.
2. Confirm the exact apparatus: material/component/lot, nozzle/mixer, pump, motor/drive, valve, load-cell/ADC/filter, temperature/pressure sensing, MCU/board, sample times, safe state, and acceptance limits. Mark missing facts `UNKNOWN`.
3. Read only the references needed:
   - `references/toolchain-selection.md` for Python, OpenModelica, Renode, Scilab/Xcos, Octave, or MATLAB/Simulink routing.
   - `references/control-theory.md` before choosing PID, dead-time compensation, run-to-run learning, observers, or MPC.
   - `references/model-and-evidence.md` for plant equations, uncertainty, Monte Carlo, SIL/PIL/HIL, and evidence labels.
   - `references/source-catalog.md` before relying on a library, product, paper, or vendor claim.
4. Copy `assets/taeha_dispense_sim` into a project-owned folder when an executable starting point is useful. Never modify the bundled template in place for a product project.

### Fidelity ceiling

Do not use this skill as a substitute for motor-current-loop design, CFD, machine safety certification, or physical accuracy evidence. Route motor/peripheral contracts to `$stm32-device-driver`, firmware safety review to `$embedded-code-audit`, and real test execution to `$embedded-test-debug`. Keep hardware claims `UNVERIFIED` until board-specific evidence exists.

## Toolchain decision

Use the smallest sufficient stack.

- Start with Python, NumPy/SciPy, and pytest/unittest for portable controller logic, parameter sweeps, captured-data replay, Monte Carlo, persistence faults, and CI.
- Add OpenModelica plus FMI when coupled electrical, rotational, thermal, hydraulic, or compliance dynamics need component-based physical modeling.
- Add Renode only for firmware-visible CPU/peripheral/interrupt behavior that its exact machine model implements.
- Use real board HIL for ADC/DMA timing, motor-drive interfaces, electrical faults, load-cell latency, brownout persistence, and safe outputs.
- Choose MATLAB/Simulink when the team needs its integrated multi-domain libraries, model-based code generation, requirements/coverage/formal verification, or supported real-time HIL. Do not choose it merely because PID is required.

## Required model layers

Keep these independently replaceable:

1. **Actuator:** command limits, deadband, quantization, slew, current/torque/speed loop, motion lag, backlash/slip, drive faults, stale command.
2. **Fluid/process:** pressure/compliance, transport delay, nozzle resistance, temperature/viscosity dependence, valve lag, cavitation/air, cutoff and dribble.
3. **Measurement:** load-cell mechanics, excitation/reference, ADC/filter group delay and settling, sample timestamps, noise, vibration, creep, drift, saturation, staleness.
4. **Control:** hard envelope, profile feed-forward, optional in-shot feedback, shot-to-shot correction, A/B ratio supervisor, confidence and rollback.
5. **Persistence:** schema, units, profile key, provenance, CRC, sequence, dual-copy/transaction, factory fallback, power-loss injection.
6. **Safety:** state machine and independently enforced current, pressure, travel, temperature, timeout, interlock, ratio, watchdog, and safe-output limits.

## Control selection order

1. Fit and validate a feed-forward map over target mass, component, material/lot, nozzle/mixer, temperature, pressure, pump, gear ratio, and actuator.
2. Identify first-order-plus-dead-time or a justified nonlinear model from experiments. Record excitation, units, residuals, validation data, and validity range.
3. Use cascaded motor current/torque, speed, and position loops inside the drive or low-level firmware; do not use final shot mass as a substitute for these fast loops.
4. Prefer bounded shot-to-shot EWMA/integral correction when the stable mass arrives after cutoff or settling.
5. Add in-shot PI/PID only when measurement bandwidth and total delay support the required closed-loop bandwidth. Use anti-windup, output/rate limits, derivative filtering, bumpless transfer, and stale-data handling.
6. Consider a Smith predictor only after delay and delay variation are identified. Fall back safely when model mismatch exceeds its validated range.
7. Use gain scheduling across validated operating regions. Keep each material/lot/nozzle/mixer/temperature-band profile isolated.
8. Use MPC, MHE/Kalman observers, RLS, or iterative learning only when simpler bounded control fails a measured requirement and compute time, constraints, robustness, rollback, and embedded verification are proven.

## Execution workflow

1. Define requirements and an evidence matrix.
2. Run the template baseline and preserve inputs, seed, configuration, outputs, plots, and hashes.
3. Replace nominal constants with measured profile data; keep calibration and validation datasets separate.
4. Sweep target, temperature, ratio, delay, drift, friction, viscosity, noise, and actuator type. Include boundary and fault cases.
5. Verify invariants with unit/property/fuzz tests: bounds, profile isolation, invalid-measurement rejection, tick wrap, fixed-point range, CRC fallback, interrupted writes, and no automatic unsafe re-enable.
6. Compare controller candidates on bias, spread, ratio error, rejected shots, convergence, actuator effort, constraint violations, and recovery—not mean error alone.
7. Run SIL, then generated/handwritten-code equivalence tests, then PIL or target build, then outputs-disabled board tests, then instrumented HIL and repeated physical shots.
8. Label results `SIMULATED`, `SIL`, `PIL`, `HIL`, or `PHYSICAL`. Label exact-board claims `UNVERIFIED` until measured.

## Release gates

- No adaptation from invalid, stale, saturated, unstable, purged, manually altered, faulted, or profile-mismatched shots.
- Every update has absolute, rate, confidence, sample-count, and validity-envelope limits plus rollback.
- A/B ratio uses independent A and B measurements or a separately validated measurement chain. Total mass alone cannot certify ratio.
- Simulation includes sensor delay, cutoff tail, quantization, saturation, noise, drift, and realistic parameter mismatch.
- HIL evidence identifies board revision, firmware hash, calibration/profile revision, material/lot/nozzle/mixer, instruments, environment, raw data, pass/fail criteria, and fault restoration.
- Physical safety remains independent of the simulation or adaptive controller.
