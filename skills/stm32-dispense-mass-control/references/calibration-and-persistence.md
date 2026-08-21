# Calibration and Persistence

## Measurement calibration

- Capture zero/tare under the actual fixture and container conditions.
- Use at least one traceable known mass spanning the operating region; a two-point fit is the minimum. Use multiple points and residual analysis when linearity is not established.
- Record ADC/reference configuration, excitation, sample rate, filter, settling rule, temperature, fixture, orientation, operator/tool, date, source masses, raw readings, fitted coefficients, residuals, and validity range.
- Characterize short-term noise, repeatability, creep, zero return, hysteresis, temperature drift, vibration sensitivity, saturation, and recovery.

## Process calibration

Calibrate per compatible profile. Sweep target mass, temperature, speed/acceleration, pressure, cutoff, precharge, decompression, settle time, and relevant nozzle/mixer conditions. Randomize or bracket runs when drift is possible. Preserve raw trials and rejected-trial reasons. A density curve converts volume to expected mass only within its measured material and temperature domain.

## Stored record

Include magic, schema version, total length, profile key/hash, units, coefficients, limits, sample count, confidence/quality indicators, timestamp or sequence, firmware compatibility, provenance, and CRC. Prefer dual slots or append/commit records:

1. Write the inactive record with an incremented sequence.
2. Read back and validate length, range, schema, and CRC.
3. Atomically mark it committed if the medium requires a commit marker.
4. On boot, choose the newest fully valid compatible record.
5. Fall back to a safe factory/default profile with compensation disabled when neither record is valid.

Rate-limit flash writes and account for endurance. Test reset or brownout at each write boundary. Never silently reinterpret an old schema or default units.

## Schema migration

Treat migration as an explicit maintenance state with outputs disabled: read the old record, validate its exact schema and provenance, convert with checked units and ranges, write a new inactive record, read back and verify, then commit it. Quarantine learned correction and require revalidation whenever semantics, units, profile keys, sensor chain, or control limits changed. Keep the old valid record or factory fallback until the new record is fully committed; never migrate implicitly in a dispense control path.
