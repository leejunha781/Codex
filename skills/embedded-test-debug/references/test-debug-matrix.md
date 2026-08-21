# Test and debug matrix

| Area | Host test | Target test | Fault test |
|---|---|---|---|
| RS-485 parser | length, CRC, fuzz, timeout wrap | DE timing, DMA, back-to-back frames | disconnect, noise, duplicate frame |
| Dispenser state | sequence, illegal event, profile isolation, bounded correction, calibration/NVM | motor/valve/load-cell timing and canonical mass matrix | sensor stuck/stale/saturated, jam, pressure/temp/ratio, brownout; apply `stm32-dispense-test-debug.md` and `$stm32-dispense-mass-control` `references/verification-matrix.md` in full |
| BLDC control | limits and stale command | PWM/timer and enable polarity | driver fault, overcurrent, comm loss |
| Stepper control | homing state and position bounds | pulse timing and direction | limit switch, jam, missed home |
| UART servo gateway | framing, CRC, replay, scaling, slew, stale timeout | DMA/IDLE, baud, response timing, enable/brake polarity | corrupt/lost stream, overflow, drive alarm, limit, emergency stop |
| Firmware reconstruction | recovered parser/state-machine differential tests | GPIO/peripheral/protocol capture comparison | brownout, watchdog, comm loss, fault parity |
| Watchdog/reset | progress model | reset reason and recovery | dead task, ISR storm, brownout |
| I2C/SPI register IC | codec, masks, reset/readback, wrong ID, timeout | bus timing/mode/address/chip select, DMA/cache, concurrency | NACK, stuck bus, truncated/status fault, reset mid-transfer |
| CAN/FDCAN motor node | ID/DLC/scaling/freshness, malformed/fuzz frames | filters, arbitration load, heartbeat, latency | bus-off, duplicate/stale command, disconnect, recovery storm |
| External motor-driver IC | configuration image, quantization, legal transitions | reset-safe pins, PWM/deadtime/break, sensing, readback | nFAULT/OCP/UVLO/OT, failed clear, watchdog reset |
| ADC motor calibration | scaling, overflow, coefficient CRC/version, states | VREFINT, fixtures, PWM trigger/skew, instrument correlation | saturation, open/short, DMA freeze, drift, wrong blob, save brownout |
| LCD GUI | model/presenter/navigation, bounds, queue policy | first frame, color/touch grid, frame/flush time, memory | missing devices, corrupt assets, queue overflow, DMA timeout, soak |

Record firmware hash, board revision, probe serial, power limits, stimulus, expected/actual result, logs, and measurement files.
