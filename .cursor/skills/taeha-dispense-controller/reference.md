# Taeha Dispense Reference

## Operation Modes (Public Controllers)

| Mode | Intent | SW watch-outs |
|------|--------|---------------|
| Time | Dispense for fixed time | Timer vs motor lag; suck-back after stop |
| Steady | Continuous while trigger held | Pressure stability; overheat/stall |
| Metering / Volume | Fixed volume | Encoder/pulse count; `K_cal`; drip |
| Interval | Timed repeats | Period jitter; residual pressure |
| Purge | Clear path | Bound duration; waste/alarm |
| Ratio (2K) | Mix ratio | Dual-axis sync; ratio error latch |

## Typical Alarm / Fault Classes

- Motor / encoder error
- Pressure out of range / sensor fail
- Material empty / remaining low
- Interlock open / E-stop
- Communication timeout / CRC
- Calibration invalid / recipe CRC fail

On fault: enter safe-state, latch code, keep last N samples for RCA.

## Recipe Channel Memory

- Store per-channel: mode, speed/volume/time, suck-back, pressure setpoints, calibration factors
- Version + CRC; reject load if invalid
- Change management: never silently mutate running channel mid-shot

## Bring-up Order (New Controller Board)

1. Power rails and reset
2. Clock / SWD
3. GPIO defaults (fail-safe: outputs deasserted)
4. UART console / LED heartbeat
5. Motor driver enable path (current limit first)
6. Sensors + ADC calibration
7. Host protocol echo
8. Closed-loop short dispense with scope on pulse/PWM

## 30 / 60 / 90 Onboarding (Controller SW)

| Window | Focus | Done when |
|--------|-------|-----------|
| 0–30d | Part number, codebase, schematics, I/O map, motor/sensor, protocol, test baseline | Env reproducible; can explain interfaces |
| 31–60d | Reproduce SM + motor/I/O + serial/Ethernet paths; close one RCA | Measured RCA + regression |
| 61–90d | Calibration, motor fault, bus integrity, recipe versioning, test automation | Data-backed improvement + team asset |

## Public Manuals (starting points)

- [PROCON-100 Operation Manual](https://dispenserobotics.com/files/Manual/ProCon-100%20Operation%20Manual_R3.7_eng.pdf)
- [PDC-100 Operation Manual](https://dispenserobotics.com/files/Manual/PDC-100%20Operation%20Manual_Rev2.1_eng.pdf)
- [TAEHA catalog (EN)](https://www.syneo-solutions.com/wp-content/uploads/2025/12/Nouveau-Catalogue_TAEHA_EN_OCT-25_compressed.pdf)

Treat manuals as **behavioral requirements**, not as internal firmware design truth.
