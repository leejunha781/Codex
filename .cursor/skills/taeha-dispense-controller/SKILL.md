---
name: taeha-dispense-controller
description: Taeha precision dispense controller firmware patterns — state machines, pump/motor control, recipe/calibration, Modbus I/O, safety interlocks. Use for Taeha, PROCON/PDC/MPCON, dispensing, suck-back, shot metering, or industrial dispenser controller work.
---

# Taeha Dispense Controller Development

## When to Use

- Dispenser controller firmware (STM32-class MCU assumed for new work; confirm part)
- Recipe, calibration, suck-back, purge, metering modes
- Pump / valve / pressure / flow / remaining-material sensing
- Host/PLC links (RS-232, RS-485 Modbus RTU, Ethernet Modbus TCP)
- Safety: E-stop, interlock, fault latch, safe-state

Public product context only — **do not invent Taeha private schematics or secret firmware architecture**. Prefer asking for internal docs when needed.

Also load `stm32-device-drivers` for MCU/driver work and `embedded-codex-workflow` for Cursor↔Codex flow.

## Product Domain (Public Sources)

Taeha: liquid precision dispensing total solution (pump, valve, controller, material supply ↔ robot/vision/inline).

| Controller (public) | Notable traits |
|---------------------|----------------|
| PROCON-100 | Time / Steady / Interval; speed/volume/suck-back/calibration; motor/encoder errors; RS-232 |
| PDC-100 | Multi-motor; Time/Steady/Purge/Ratio; pressure/flow/remaining; RS-485 Modbus RTU |
| MPCON | Ethernet Modbus TCP; RS-232/485 Modbus RTU |
| TP-50 | Peristaltic; microstep; low-viscosity airless fluids |

Quality axes for controller SW: **precision, repeatability, response time, recipe reproducibility, alarm/interlock, diagnostics, uptime**.

## Canonical State Machine

```
INIT → IDLE → READY → DISPENSE → SUCK_BACK → COMPLETE
                 ↘ PURGE | CALIBRATE
FAULT (latched) — separate path; never auto-clear on reset alone
```

**Guards before DISPENSE:** recipe valid, material present, interlock closed, sensors plausible, motor ready.

**Safety priority:** E-stop / overcurrent / limit > communication > host command.

**Safe-state:** disable PWM/pulse, close valve, reject commands, latch fault, preserve evidence.

## Volumetric Control Model

Treat shot quality as a **system** problem:

```
Motor rotation → Pump displacement → Pressure/flow feedback → Suck-back timing
```

- Theoretical: `V ≈ K_cal × N` (or `Q ≈ K_cal × ω`)
- Real `K_cal` verified under material, temperature, pressure, nozzle conditions
- 2K: track cumulative volume per axis, ratio error, sync error together

## Firmware Module Map

| Module | Responsibility |
|--------|----------------|
| `app_sm` | State machine, guards, transitions |
| `recipe` | Channel memory, versioning, CRC/validate |
| `motor_svc` | DC/BLDC/STEP/SERVO setpoints + faults |
| `sensor_svc` | Pressure, flow, remaining, plausibility |
| `io_svc` | Digital I/O, interlock, dosing-end signal |
| `host_proto` | Modbus RTU/TCP or proprietary serial |
| `diag` | Alarm codes, ring log, last-shot metrics |

Keep STM32 HAL behind `stm32-device-drivers` layering.

## Host / I/O Expectations

- External start: contact or NPN OC
- Dosing end: NPN OC (typical public manuals)
- Serial: RS-232 point-to-point; RS-485 2-wire with DE/RE discipline
- Fieldbus: Modbus RTU/TCP where product family requires it
- Never block the control loop on host TX; use queues

## Verification Ladder

1. Unit (SM transitions, CRC, parsers)
2. SIL / driver smoke
3. Bench with pump + sensors
4. HIL / robot handshake
5. Fault injection (loss of material, encoder fault, bus drop)
6. Temperature / voltage / EMI where applicable
7. Endurance + power-cycle

**Metrics:** shot mean/σ/Cpk, cycle time, ratio error, alarm/recovery, packet error, motor current/temperature.

## Accuracy Boundaries (Joonha Context)

Do **not** claim past ownership of: Taeha production dispenser/STEP products, STM32 Ethernet stack from scratch, Modbus greenfield, FreeRTOS product deploy, or FOC/PID algorithm authorship — unless the current task actually implements them.

Credible anchors: NXP LPC1769 KSS-III firmware+I/O+RS-485/232 full lifecycle; Intellian STM32F BLDC code validation/tuning; RM57L843 EtherCAT/servo **research** (not production).

## Notion Hub

Internal prep (user-owned): [태하 제어기 개발(SW) 기술면접 준비 허브](https://app.notion.com/p/3aa2392cd0dc8195be19ef0f3971b59b)

## Related

- [reference.md](reference.md) — modes, alarms, 30/60/90 onboarding
- `stm32-device-drivers`
- `embedded-codex-workflow`
