# STM32 Dispense Audit

Apply after `$stm32-dispense-mass-control` and the exact device-driver contracts are loaded.

## Traceability and separation

- Exact STM32/board/schematic/HAL/LL/RTOS/toolchain and build configuration are recorded.
- Portable estimator, planner, compensator, ratio supervisor, and persistence logic do not directly own HAL registers or blocking I/O.
- Motor commands, valve commands, measured mass, estimated flow, and two-component ratio use explicit units and independent validity states.

## Sensor and timing

- ADC/SPI/DFSDM sample timing, digital-filter delay, DMA/cache coherency, timestamp source, stale threshold, saturation, overflow, dropped samples, and task/ISR ownership are correct.
- Stability and outlier checks cannot accept NaN/overflow/wrapped values or silently reuse stale data.
- Feedback bandwidth is justified against measured transport and actuator delay; delayed measurements cannot destabilize in-shot control.

## Compensation and safety

- Updates require accepted shots, matching profile/calibration schema, minimum samples, valid target range, and healthy sensors.
- Gain, integral, error, accumulated correction, per-shot delta, total correction, motor output, pressure, current, temperature, travel, and ratio are bounded with proven arithmetic.
- Limit saturation triggers diagnostics and freeze/fault behavior; code cannot expand limits or learn around faults.
- Profile changes, purge/prime/manual shots, maintenance, sensor invalidation, repeated outliers, reset, or incompatible firmware quarantine learned state.
- Startup, disabled, fault, watchdog, brownout, comms loss, sensor loss, jam, blockage, and empty-material states produce defined safe outputs and controlled recovery.

## Calibration and NVM

- Raw counts are not exposed as mass without valid tare/calibration and units.
- Record includes magic, schema, length, profile, units, coefficients, limits, provenance/sequence, and CRC.
- Interrupted writes cannot replace the last valid record; boot selection and factory fallback are deterministic.
- Flash endurance and write-rate limits are enforced; malformed/old/incompatible records fail closed.

## Two-component control

- Total mass is not used as proof of A/B ratio.
- Individual component estimates, density versus temperature, startup lag, purge, pot life, mixer pressure, and ratio fault window are explicit or labeled `UNVERIFIED`.
