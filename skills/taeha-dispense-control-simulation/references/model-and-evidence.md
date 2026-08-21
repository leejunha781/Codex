# Model and evidence workflow

## Minimum plant model

Represent each component with a command path, actuator lag, transport delay, fluid/compliance lag, and cutoff tail. A useful first model is:

[
\tau_a \dot a + a = sat(rateLimit(u))
]
[
\tau_f \dot q + q = K(T,material,nozzle) a(t-L)
]
[
\dot m = max(0,q+q_{dribble})
]
[
\tau_d \dot q_{dribble}+q_{dribble}=0
]

Add deadband, step quantization, backlash/slip, valve delay, pressure dependence, cavitation/air, and non-Newtonian behavior only when data shows they matter.

For temperature, start with a profile-specific empirical factor. Do not claim a universal viscosity law. Use measured rheology or vendor data over the validated range.

## Identification experiment

1. Confirm safe limits and instrument calibration.
2. Excite one variable at a time where possible: motor speed/steps, valve timing, temperature, pressure, or target.
3. Record commanded and actual actuator state, current/torque, pressure, flow if available, raw ADC, filtered mass, temperature, timestamps, and faults.
4. Synchronize clocks or estimate offset.
5. Reserve separate calibration and validation datasets.
6. Fit FOPDT, state-space, Hammerstein/Wiener, or nonlinear parameters appropriate to observed behavior.
7. Inspect residual bias, autocorrelation, heteroscedasticity, outliers, and parameter confidence.
8. Reject a model that only fits the calibration sequence.

## Uncertainty and Monte Carlo

Sweep or sample:

- material lot and age;
- density and viscosity versus temperature;
- nozzle/mixer restriction and wear;
- supply pressure and depletion;
- motor gain, friction, backlash, slip, missed steps;
- valve delay and cutoff tail;
- ADC reference, gain, noise, drift, filter delay, sample loss;
- task jitter, communication delay, queue overflow;
- profile mismatch and corrupted persistence.

Report percentile and worst observed performance plus constraint violations. A narrow mean error is not enough.

## Verification ladder

| Level | What it can prove | What remains unverified |
|---|---|---|
| Model simulation | equations and controller behavior under declared assumptions | plant fidelity and hardware |
| Host/SIL | portable production logic, bounds, fixed-point, persistence, captured replay | compiler target, HAL, ISR, electrical timing |
| Virtual target | modeled CPU/peripheral/interrupt behavior | unmodeled peripheral details, analog, motor power, load cell, physical safety |
| PIL/target build | generated/handwritten code on target processor and timing samples | closed-loop apparatus behavior |
| HIL | firmware plus electrical interfaces and injected plant/fault model | exact material/mechanics unless included physically |
| Physical repeated shots | apparatus accuracy/repeatability under tested conditions | untested profiles, lots, environments, wear, life |

## Required automated tests

- target and ratio envelope boundaries;
- all actuator types and command saturation;
- temperature/nozzle/profile selection;
- delay zero, nominal, and worst validated;
- stale, frozen, noisy, saturated, NaN/invalid sensor;
- prime/purge/manual/fault shots do not learn;
- profile mismatch does not read or update another profile;
- absolute/rate/sample/confidence learning bounds;
- CRC bit error, both-slot corruption, stale sequence, wrong schema/profile, reset at every write phase;
- timeout, queue overflow, watchdog reset, communication loss;
- A/B startup lag, one-side empty/blockage/slip, ratio alarm and latched safe response;
- deterministic seed and reproducible artifact manifest.

## HIL evidence record

Capture:

- requirements and pass/fail thresholds;
- board/fixture/instrument revisions and calibration status;
- firmware ELF/BIN hash, build options, tool versions, profile/calibration schema and sequence;
- material/component/lot, age, temperature, humidity, nozzle, mixer, supply pressure;
- raw and filtered sensor streams with timestamps;
- commands, motor feedback, valve, pressure, current, faults and state transitions;
- injected fault and restoration procedure;
- bias, standard deviation, confidence interval, rejected shots, ratio error and constraint violations;
- reviewer and disposition.

Keep outputs disabled until polarity, limits, safe state, and recovery are confirmed. Use current limiting and containment appropriate to the material and power stage.
