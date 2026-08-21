# UART input to servo-drive architecture

Use the drive vendor's protocol and electrical manual as the authority. Treat Modbus RTU, ASCII commands, proprietary binary frames, pulse/direction, and analog commands as different adapters; never infer one from another.

## Required contract

Record both UART sides when the controller receives values on one link and commands a drive on another:

- Peripheral, pins, electrical standard, transceiver, isolation, termination, baud, word length, parity, stop bits, inversion, and duplex.
- Frame delimiter or length, byte order, address, command/register, payload type, units, scaling, CRC/checksum, response, timeout, retry, duplicate/sequence handling, and exception mapping.
- Position, velocity, torque/current, acceleration, deceleration, jerk, software/travel limits, homing, enable, brake, alarm reset, emergency stop, and stale-command timeout.
- ISR/DMA ownership, ring-buffer capacity, overflow behavior, command mailbox ownership, and diagnostic counters.

## Source layout

Keep generated/CubeMX code and product logic separate:

```text
App/
  motor_command.[ch]       normalized fixed-width commands and units
  command_validator.[ch]   ranges, slew limits, sequence and freshness
  safety_supervisor.[ch]   interlocks, stale timeout and fault latch
Protocol/
  input_frame.[ch]         inbound UART framing and CRC
  servo_protocol.[ch]      vendor-specific drive frames/registers
Platform/
  uart_rx_dma.[ch]         DMA/IDLE adapter and bounded event handoff
  servo_transport.[ch]     TX/RX turnaround, timeout and retry
  board_safety.[ch]        enable, brake, alarm and emergency I/O
Diagnostics/
  motor_diag.[ch]          counters, last fault and capture hooks
Tests/
  test_input_frame.c
  test_command_validator.c
  test_safety_supervisor.c
  test_servo_protocol.c
EWARM/
  project.ewp, project.eww, device.icf
```

## Data and control flow

1. DMA writes only to an owned RX buffer. ISR/IDLE handling snapshots indices and posts bounded work; it does not parse, format logs, wait, or command the drive.
2. The parser validates delimiter/length/address/CRC before decoding fixed-width fields.
3. Normalize fields into explicit units such as `_mdeg`, `_mrpm`, `_mnm`, `_ma`, and `_ms`. Reject NaN-like encodings, reserved modes, overflow, replay, and impossible transitions.
4. Validate absolute limits, rate-of-change, command age, machine interlocks, homing state, and drive readiness.
5. Publish one complete command atomically to the control owner. The vendor adapter serializes it and waits only with a bounded timeout outside ISR context.
6. The safety supervisor can preempt the pipeline and request controlled stop or immediate disable according to the documented hazard response.

## State model

Use explicit states such as `SAFE_DISABLED`, `INITIALIZING`, `READY`, `ENABLED`, `CONTROLLED_STOP`, and `FAULT_LATCHED`. Define legal events, entry outputs, deadlines, and recovery for each state. Do not clear a latched drive or safety fault merely because a new UART command arrived.

## Verification

- Host: frame boundaries, CRC, malformed/fuzz input, integer scaling, tick wrap, sequence/replay, ranges, slew, stale timeout, and every state transition.
- Target: baud tolerance, DMA wrap/overrun, IDLE timing, back-to-back frames, DE timing where applicable, response latency, and watchdog progress.
- HIL: unplugged UART, corrupt frame, duplicate frame, drive alarm, limit input, emergency stop, brownout/reset, stalled motor, and command stream loss.

Capture expected and actual bytes, timestamps, board/drive revisions, parameter set, firmware hash, power/current limit, and safe-state output levels.
