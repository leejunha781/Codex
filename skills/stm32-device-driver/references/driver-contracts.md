# Driver contracts

For every driver, record hardware part/revision, bus instance, pins, polarity, clock, electrical limits, API, units, ownership, timing, concurrency, buffer lifetime, errors, recovery, diagnostics, and tests.

## RS-485 checklist

- UART format and baud tolerance
- DE/RE polarity and setup/hold timing
- DMA/interrupt ownership
- inter-byte and frame timeout
- maximum payload before indexing
- CRC coverage and byte order
- retry/backoff and duplicate-command policy
- bus-idle and collision behavior

## Motor checklist

- safe output state during reset/init
- enable/brake/direction polarity
- speed/position/current limits
- command freshness timeout
- end-stop, homing, stall/jam, and driver fault handling
- acceleration/deceleration constraints
- fault latch and explicit recovery
