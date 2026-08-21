# Taeha controller software profile

## Confirmed context

- Company: Taeha, dispenser and automation equipment, NANO CENTER.
- Role: controller development (software), corporate R&D.
- Strong alignment: STM32, RS-485, RS-232, BLDC, motor-drive integration.
- Candidate evidence: STM32F303 firmware, RS-485 station bus, RS-232 monitor/VFD, motor-drive unit, STM32F BLDC verification/tuning.
- Gaps to close with projects and tests: dispenser sequencing and stepper/STEP control.

## Engineering emphasis

Prioritize deterministic sequencing, serial communications, motor safety, field diagnostics, board bring-up, production testability, and maintainable Embedded C. Treat the specific MCU, board, RTOS, IDE version, protocol frame, motor driver, and pin mapping as unconfirmed until source artifacts are supplied.

## Portfolio tasks

1. RS-485 framed transport with CRC, timeout, retry, and diagnostics.
2. Dispenser sequence state machine with interlocks and fault recovery.
3. BLDC command/safety adapter with stale-command timeout.
4. Stepper axis controller with homing, limits, acceleration constraints, and jam detection.
5. Host tests plus target/HIL verification matrix.
