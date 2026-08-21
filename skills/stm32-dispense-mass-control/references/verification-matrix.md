# Verification Matrix

This is the canonical verification matrix for `$stm32-dispense-mass-control`. `embedded-test-debug/references/stm32-dispense-test-debug.md` adds debugging observability but must not weaken or replace this matrix.

## Static and supply-chain evidence

- Exact production compiler warnings plus the project-selected static analysis; run host sanitizers where portable modules permit.
- Record tool version and configuration, suppressions and owners, reports, dependency tag/commit/license/notices/security status, and SPDX or CycloneDX SBOM impact.

## Host

- Unit tests: profile matching, unit conversion, filters, stability detector, shot acceptance, bounded correction, saturation, rollback, state transitions, fault latching, and NVM selection.
- Property/fuzz tests: parsers and calibration blobs never escape limits; corrupt length/CRC/version is rejected; arbitrary sensor sequences never update learning when invalid.
- Boundary tests: zero/min/max target, negative/raw overflow, divide-by-zero guards, fixed-point extremes, timer wrap, sequence wrap, stale sample, repeated outlier, and correction limit hits.
- Replay/model tests: captured mass/current/pressure/temp traces, transport delay, drift, sensor noise, residual dribble, and missed-step scenarios.

## Target

- Cross build and link map with exact compiler/options; stack and memory budget.
- Reset/brownout outputs disabled; watchdog and fault paths; ISR/task latency and load.
- ADC/SPI/DFSDM data-ready timing, DMA/cache coherency, timestamp accuracy, digital-filter settling, and sample-loss behavior.
- Motor/valve waveform, direction/polarity, current/pressure/travel limits, comms timeout, and controlled stop measured with appropriate instruments.
- Calibration-store interrupted-write and endurance policy.
- No flash erase/program occurs in a deadline-critical control context; bounded deferred persistence and queue-full/failure behavior are exercised.
- For cache-less parts prove DMA ownership and ordering. For F7/H7 or other cached parts verify linker/MPU placement, cache-line alignment, clean/invalidate direction, and barriers using the production cache configuration.

## HIL and physical process

- Load-cell simulator plus traceable weights; unplug, stale, saturation, overload, noise, and drift faults.
- Temperature sweep, pressure/backpressure change, nozzle restriction, air bubble, empty/low supply, jam, valve delay, motor slip, and comms loss.
- Material/lot/nozzle/profile change proves learned-state quarantine.
- Two-component startup, purge, individual A/B shortage, ratio deviation, pot-life expiry, and mixer-pressure fault.
- Repeated shots across the validated operating space. Report target, actual mass, error, bias, standard deviation, coefficient of variation, confidence interval, rejection rate, conditions, firmware hash, board and calibration revision. Use Cp/Cpk only after the process is shown stable and the specification is legitimate.

Never convert successful simulation or host tests into a hardware accuracy claim.
