# Toolchain selection

## Decision matrix

| Stack | Best use | Strengths | Important limits |
|---|---|---|---|
| Python + NumPy/SciPy + Matplotlib | First executable model, identification, Monte Carlo, captured replay, CI | Open, scriptable, testable, close to portable C logic | Build physical libraries and embedded timing evidence yourself |
| python-control | LTI/state-space analysis, frequency response, classical and optimization-based control | Familiar control API; can use SciPy optimization | Not a plant library or HIL system |
| CasADi/acados | Constrained nonlinear optimization, MPC/MHE, code generation | Algorithmic differentiation and embedded-oriented solvers | Requires model/solver expertise and worst-case-time proof |
| OpenModelica + Modelica Standard Library | Electrical–mechanical–thermal–fluid component modeling | Open, acausal, reusable multi-domain components, FMI exchange | Model fidelity and real-time suitability must be validated |
| Scilab/Xcos | Free block-diagram teaching and classical control | Visual continuous/discrete/hybrid simulation, PID and delay blocks | Smaller embedded verification/codegen ecosystem than Simulink |
| GNU Octave control package | MATLAB-like scripts, classical control education | Free and low switching cost | No Simscape-equivalent integrated physical/HIL stack |
| Renode | Firmware CPU/peripheral/interrupt integration and automated virtual-target tests | Deterministic automation, custom peripheral models, Robot Framework | Does not prove unmodeled analog, motor power stage, load cell, DMA/cache timing, or physical safety |
| Native host tests/SIL | Portable state machines, codecs, fixed-point, persistence | Fast and CI-friendly | HAL/ISR/electrical behavior is mocked |
| Board PIL/HIL | Real firmware, interfaces, faults, timing | Essential target evidence | Needs safe fixture, instruments, calibration, and reproducible automation |
| MATLAB/Simulink family | Integrated physical model, control design, codegen, requirements and commercial HIL | Mature end-to-end model-based workflow and vendor support | License cost, tool qualification/configuration, and model-code equivalence work remain |

## Recommended Taeha stack

Use a layered hybrid rather than one universal simulator:

1. Python reference model and automated tests for every controller change.
2. OpenModelica/FMI only when the pump, compliance, thermal, electrical, and mechanical interactions justify the added detail.
3. Renode for supported MCU/peripheral behaviors and firmware integration.
4. Real STM32 board HIL for sensor/actuator timing, faults, persistence, safe state, and production build.
5. MATLAB/Simulink when commercial workflow benefits outweigh licensing and migration cost.

## When MATLAB is the strongest choice

Choose MATLAB/Simulink when several of these are binding requirements:

- Simscape Fluids and Simscape Electrical must be coupled with controller and state-machine models.
- The team requires Control System Toolbox, Simulink Control Design, System Identification Toolbox, or Model Predictive Control Toolbox in one supported environment.
- Production uses Embedded Coder, Fixed-Point Designer, Stateflow, Simulink Test, Requirements Toolbox, Simulink Coverage, Design Verifier, or Polyspace.
- A supported Simulink Real-Time/Speedgoat HIL route is already available.
- Suppliers already exchange Simulink models and the organization can govern model versions and licenses.

A practical MATLAB package for this dispenser would usually be:

- MATLAB + Simulink
- Simscape + Simscape Fluids
- Simscape Electrical for DC/BLDC/PMSM/ACIM and drive models
- Control System Toolbox
- Simulink Control Design
- System Identification Toolbox
- Stateflow
- Fixed-Point Designer
- Embedded Coder
- Simulink Test and Simulink Coverage
- Requirements Toolbox
- Simulink Design Verifier or Polyspace when assurance goals justify them
- Simulink Real-Time plus target hardware for commercial HIL
- Motor Control Blockset when the low-level PMSM/BLDC/ACIM drive itself is modeled and deployed
- Model Predictive Control Toolbox only if constrained MPC is actually selected

Do not buy the full list by default. Start from required evidence and acceptance criteria, then confirm current license dependencies with MathWorks.
