# STM32 firmware reconstruction

Analyze only firmware the user is authorized to inspect and reimplement. Preserve original artifacts read-only and record SHA-256 hashes.

## Evidence limits

| Artifact | Usually establishes | Does not reliably establish |
|---|---|---|
| PDF schematic and BOM | MCU, nets, pull states, transceivers, drive ICs, nominal polarity | Runtime timing, protocol meaning, calibration, recovery policy |
| Intel HEX or BIN | Vector table, code/constants, register access, interrupt use, observable algorithms | Original C, names, comments, types, modules, exact compiler options |
| Motor specification | Ratings, speed/torque envelope, feedback and thermal limits | External drive command interface unless explicitly included |
| Servo-drive manual | Command levels, frames/registers, enable/brake/alarm, timing and limits | Product sequencing and application-specific recovery |

Request matching `.out`/ELF with debug data, `.map`, `.ewp`, `.eww`, `.icf`, `.ioc`, BOM, external-IC manuals, protocol documents, parameter dumps, communication captures, and a golden board. Record missing items as `UNKNOWN`; never silently replace them with assumptions.

## Procedure

1. Confirm authorization, board revision, MCU order code, firmware source, flash base, option bytes/readout protection, and artifact hashes.
2. Extract PDF text and vector geometry when available. Render scanned pages and use OCR only as a candidate source. Manually verify every MCU pin, net label, polarity, connector, voltage, transceiver, motor-drive signal, and sheet cross-reference against the page image.
3. Parse HEX record checksums, address extensions, gaps, entry records, initial stack pointer, and reset vector. Validate addresses against the exact MCU datasheet and memory map.
4. Load the image as little-endian Arm Thumb code for the confirmed core. Define vector table, memory regions, peripheral addresses, literal pools, startup/runtime routines, and interrupt handlers.
5. Cross-reference register accesses with the exact ST reference manual and errata. Build pin, peripheral, DMA, interrupt, clock, and external-device tables with evidence and confidence.
6. Identify state machines, frames, CRC/checksum logic, lookup tables, timing constants, fault paths, and compiler-library idioms. Treat decompiler output as evidence, never recovered source.
7. Create a clean IAR EWARM project with the exact device, startup, CMSIS/HAL or LL version, `.icf`, FPU/ABI, runtime library, stack/heap, optimization, language settings, and linker map generation.
8. Express observed behavior as requirements and tests, then implement portable domain logic separately from STM32 adapters.
9. Compare original and reconstructed behavior using reset/clock/GPIO measurements, UART or RS-485 captures, PWM, ADC scaling, motor commands, watchdog, brownout, and injected faults.

## Equivalence levels

- L0: IAR project builds and links for the confirmed MCU.
- L1: reset, clock, GPIO, interrupt, DMA, and peripheral initialization match measurements.
- L2: communication framing, timing, CRC, sequencing, retries, and diagnostics match captures.
- L3: servo/motor sequencing and limits match nominal and boundary tests.
- L4: watchdog, brownout, jam/stall, limit, communication-loss, and drive-fault behavior reach the documented safe state.

Do not claim completion above the highest measured level. Byte-identical output is not an acceptance criterion unless the original sources, libraries, toolchain version, and build options are available.

## Authoritative documents

Use the exact STM32 datasheet, reference manual, programming manual, errata, and application notes from ST; the matching Cortex-M architecture documentation from Arm; and the installed IAR-version documentation. Record document number and revision in the evidence matrix.
