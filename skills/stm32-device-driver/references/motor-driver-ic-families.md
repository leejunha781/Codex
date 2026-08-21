# External motor-driver IC families

Use this reference to select and implement a driver for brushed DC, BLDC/PMSM, AC induction, stepper, or an external servo drive. The motor class, power stage, command interface, feedback topology, and safety function are separate decisions. Never generalize a register map or fault sequence across suffixes in one IC family.

## Selection matrix

| Class | Representative devices/projects | Typical MCU interface | Driver responsibilities |
|---|---|---|---|
| Brushed DC H-bridge | TI DRV8840/DRV887x; ST VNH/L620x families | PWM+GPIO, fault GPIO, sometimes SPI | direction, PWM/decay, current limit, fault latch, braking/coasting |
| BLDC/PMSM gate driver | TI DRV8323/DRV83xx; STSPIN32 families | 1/3/6 PWM, SPI or hardware straps, ADC, fault GPIO | gate configuration, CSA gain, deadtime ownership, fault decode, safe enable; MCU performs commutation/FOC unless stated otherwise |
| Stepper integrated driver/controller | ST L6470/L647x; TI DRV8711; ADI Trinamic TMC5160 | SPI, STEP/DIR, UART on selected parts, fault/status GPIO | motion/register commands, chopper/current settings, ramp state, stall/fault status |
| External servo/VFD | vendor-specific drive | RS-485/Modbus, CAN/CANopen, EtherCAT, pulse/direction, analog, enable/brake/alarm I/O | protocol adapter, scaling, state machine, stale-command timeout; drive owns the power stage |
| Open reference controller | ST MCSDK examples, SimpleFOC, moteus, VESC, selected ODrive releases | board-specific | architectural evidence for sensing, FOC, calibration, diagnostics, and tests; not a drop-in board driver |

## Implementation workflow

1. Confirm motor data: topology, rated/peak voltage and current, phase resistance/inductance, pole pairs or step angle, encoder/Hall/thermal interfaces, mechanical limits, and allowable stop behavior.
2. Confirm the exact driver suffix and board: power range, logic voltage, PWM mode, SPI/I2C availability, current-sense amplifier, sense resistor, gate resistors, MOSFETs, bootstrap/charge pump, fault pins, and thermal design.
3. Build a register-and-pin contract directly from the datasheet. Record reset values, write order, wake delays, locked bits, write-one-to-clear/read-clear behavior, diagnostic latching, and recovery requirements.
4. Implement `probe`, `reset`, `configure`, `verify`, `arm`, `command`, `read_status`, `clear_fault`, and `safe_disable` as legal state transitions. `clear_fault` must not automatically re-enable power.
5. Calculate current/gain/timing values in checked integer or explicitly justified floating-point code. Saturate before conversion and return the applied value when quantization matters.
6. Keep PWM generation, ADC sampling, control law, IC configuration, and machine-level safety in separate modules with explicit ownership.

## Safety rules

- Configure enable/sleep/reset and timer break inputs so reset and initialization leave power outputs disabled.
- Read and preserve the first fault snapshot before attempting recovery. Distinguish UVLO, OCP, gate-driver fault, overtemperature warning/shutdown, open load, stall, and communication failure.
- Verify PWM polarity, complementary output, deadtime, break behavior, and timer idle state with the power stage disabled before gate drive.
- Bound current, duty, torque, speed, acceleration, jerk, and command age. Apply board and motor limits below IC absolute maximums.
- For ACIM/PMSM/BLDC identification, use a documented low-energy procedure and independent overcurrent protection. Calibration is not authorization to spin.

## Authoritative device entry points

- TI DRV8323: https://www.ti.com/product/DRV8323
- TI DRV8711: https://www.ti.com/product/DRV8711
- ST L6470: https://www.st.com/en/motor-drivers/l6470.html
- ST motor drivers: https://www.st.com/en/motor-drivers.html
- ADI TMC5160: https://www.analog.com/en/products/tmc5160.html
- STM32 motor-control ecosystem and MCSDK: https://www.st.com/content/st_com/en/ecosystems/stm32-motor-control-ecosystem.html

At use time, open the exact datasheet revision, errata, evaluation-board schematic, and application notes from the product page. Do not use this catalog as a substitute for them.

## Open-source evidence

- SimpleFOC: https://github.com/simplefoc/Arduino-FOC — approachable cross-platform FOC patterns; inspect target support and license at the pinned revision.
- moteus: https://github.com/mjbots/moteus — STM32G4 servo firmware/hardware, calibration, CAN-FD, and high-rate control; Apache-2.0 unless a file says otherwise.
- VESC firmware: https://github.com/vedderb/bldc — mature STM32 BLDC/FOC and diagnostics; GPL-3.0, so treat as study material unless the product's licensing is compatible.
- ODrive: https://github.com/odriverobotics/ODrive — historical/current branches differ in hardware, licensing, and product support; pin and inspect before use.

Community projects reveal practical failure modes but do not prove electrical compatibility, functional safety, or production quality for another board.
