# STM32 Dispense Test and Debug

The canonical pass/fail matrix is `$stm32-dispense-mass-control` `references/verification-matrix.md`. This file adds debug instrumentation and execution detail; apply both, and update the canonical matrix first when requirements change.

## Host evidence

- Unit/property tests for unit conversions, filters, scale stability, outliers, profile keys, shot acceptance, state transitions, bounded correction, saturation, rollback, and NVM selection.
- Fuzz calibration/protocol blobs and arbitrary sensor sequences. Invalid or stale input must never update learned state or escape hard limits.
- Test zero/min/max target, timer wrap, fixed-point extremes, divide guards, correction saturation, repeated rejected shots, and incompatible profiles.
- Replay captured mass/current/pressure/temperature traces with measured delays; keep synthetic and measured datasets labeled separately.

## Target evidence

- Exact build and link map; stack watermark; watchdog; reset reason; ISR/task load and jitter.
- ADC/SPI/DFSDM data-ready and filter-settling timing; DMA/cache correctness; missing-sample and stale-sample behavior.
- Motor/valve direction, polarity, acceleration, current, pressure, travel, cutoff, decompression, comms timeout, and safe stop measured with outputs current/pressure limited.
- Brownout/reset at every calibration-store write phase; corrupted newest record falls back to the last valid compatible record.

## HIL and process evidence

- Inject load-cell disconnect, overload/saturation, drift, vibration/noise, unstable platform, jam, blockage, empty supply, air bubble, valve delay, motor slip, overpressure, overtemperature, ratio deviation, comms loss, watchdog, and brownout.
- Sweep validated target mass, material/lot, nozzle/mixer, temperature and pressure. Prove profile changes quarantine learned correction.
- Run enough repeated shots to report raw observations, bias, standard deviation, coefficient of variation, confidence interval, rejection rate, environmental conditions, firmware hash, board revision, and calibration/profile revision.
- For two-component material, measure A and B independently where possible; test startup transient, purge, component shortage, density/temperature change, pot-life expiry, and mixer-pressure fault.

Debug from synchronized timestamps across command, motor/valve, current, pressure, temperature, ADC raw value, filtered mass, stability decision, shot acceptance, correction state, and fault state. Never hide rejected trials.
