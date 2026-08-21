# Control Architecture

## Profile key and observability

Key control state by material/component, supplier lot, pump/head, gear ratio, nozzle or mixer, temperature band, pressure regime, and calibration schema. Store the profile identifier with every result. Do not infer unmeasured viscosity from motor current alone; current, pressure, position, temperature, and mass are complementary signals with different delays and failure modes.

## Shot sequence

1. Verify interlocks, profile identity, sensor health, tare stability, available travel/material, motor/drive readiness, and output-safe baseline.
2. Prime or pre-pressure only under a separately bounded state. Exclude prime and purge shots from learning.
3. Execute a jerk/acceleration/speed-limited motor and valve profile from the feed-forward map.
4. Apply cutoff early enough for measured transport delay and residual flow. Use bounded decompression or suck-back only if the actuator and material validation support it.
5. Wait for both process settling and scale stability; time alone is not sufficient.
6. Accept or reject the measured shot using explicit quality predicates.
7. Update the bounded compensator only for accepted, matching-profile shots. Persist according to write-rate and transactional rules.

## Compensation states

Use a state machine such as `DISABLED`, `COLLECTING`, `ACTIVE`, `FROZEN`, and `INVALID`. Activation needs a minimum accepted sample count and validated error statistics. Freeze on marginal sensor/process health. Invalidate on profile mismatch, calibration change, maintenance, incompatible firmware schema, or repeated limit hits.

Use a feed-forward table or model for the dominant behavior, then a small bounded correction. A correction that repeatedly saturates is a diagnostic that the profile, mechanism, material, sensor, or model is wrong; it is not permission to expand limits automatically.

## Timing and numeric design

- Timestamp sensor samples and commands from one monotonic clock.
- Measure ADC group delay, digital-filter latency, task jitter, actuator latency, pressure settling, and material transport delay.
- Choose fixed-point scale and accumulator width from worst-case proofs. Test negative values, zero target, minimum/maximum targets, wraparound, division guards, saturation, and unit conversion.
- Separate control-cycle deadlines from user-interface or logging work.

## Fault policy

Define detection, immediate output action, latched data, recovery authority, and diagnostic evidence for sensor stale/saturation/noise, motor driver fault, overcurrent, overpressure, overtemperature, blocked nozzle, empty supply, excessive ratio error, unexpected motion, comms loss, watchdog, brownout, NVM corruption, and repeated correction saturation.
