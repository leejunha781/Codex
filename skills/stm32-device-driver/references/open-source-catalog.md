# Open-source evidence catalog

For dispensing, load-cell, material-profile, or shot-mass-control work, also read `$stm32-dispense-mass-control` `references/evidence-catalog.md`. Apply the shared adoption record in `stm32-collaborative-development/references/open-source-qualification.md`; an entry in either catalog is evidence, not approval.

Use open source as reviewed evidence or a version-pinned dependency, not as unbounded training data. Before copying or linking code, record repository URL, tag or commit, file path, license, notices, target MCU/toolchain compatibility, and the specific pattern adopted. Recheck the current license at use time.

| Project | Appropriate use | License boundary |
|---|---|---|
| `STMicroelectronics/STM32CubeF7` and matching STM32Cube family | Vendor HAL/LL, CMSIS device, startup, IAR examples | Component licenses are mixed; inspect repository and component `LICENSE.md`. HAL is commonly BSD-3-Clause, while middleware can differ. |
| `ARM-software/CMSIS_5` | Cortex-M core access, compiler abstraction, startup conventions | Apache-2.0; prefer the device/vendor-supported CMSIS version. |
| `FreeRTOS/FreeRTOS-Kernel` | Queue/task/stream-buffer patterns and pinned RTOS dependency | MIT; verify port, interrupt-priority rules, and IAR support for the selected release. |
| `zephyrproject-rtos/zephyr` | Reference designs for UART, DMA, state machines, tests, and device models | Apache-2.0; use as architectural evidence unless the product adopts Zephyr. |
| `libopencm3/libopencm3` | Independent register-level comparison against vendor manuals | LGPL-3.0-or-later; do not paste into a closed product without a license-compatible integration decision. |
| `NationalSecurityAgency/ghidra` | Authorized disassembly, decompilation, call graphs, and scripts | Apache-2.0; decompiler output remains an inference, not original source. |
| `bialix/intelhex` | Independent validation or conversion of Intel HEX | BSD; the bundled inspector intentionally has no runtime dependency. |
| `jsvine/pdfplumber` | Machine-generated PDF text, line, rectangle, and table extraction | MIT; scanned schematics still require rendering/OCR and manual visual verification. |
| `ThrowTheSwitch/Unity` | Portable embedded-C host unit tests | MIT; isolate HAL through fakes/adapters. |
| `STMicroelectronics/STMems_Standard_C_drivers` | Platform-independent I2C/SPI register-driver patterns | BSD-3-Clause at current upstream; verify pinned files and notices. |
| `lvgl/lvgl` | Portable embedded GUI and STM32 display/input patterns | MIT; pin a compatible release and prove target memory/cache/timing. |
| `simplefoc/Arduino-FOC` | Cross-platform BLDC/stepper FOC and current-sense patterns | MIT; Arduino abstractions do not prove target timing or protection. |
| `mjbots/moteus` | STM32G4 servo control, calibration, CAN-FD, and hardware/firmware co-design | Apache-2.0 unless a file says otherwise; high-power safety remains product work. |
| `vedderb/bldc` | VESC STM32 BLDC/FOC, sensing, diagnostics, and fault patterns | GPL-3.0; do not copy into incompatible closed firmware. |
| `odriverobotics/ODrive` | Motor calibration, axis state, protocol, and historical STM32 evidence | Branches and licenses vary; pin and inspect every adopted file. |
| `mcu-tools/mcuboot` | Secure image, signing, swap/revert, and simulator patterns | Apache-2.0; target flash, key custody, rollback, and power-loss proof remain product work. |
| `TrustedFirmware-M/trusted-firmware-m` | PSA isolation and secure/non-secure boundary evidence | BSD-3-Clause core with separately licensed dependencies. |
| `Mbed-TLS/mbedtls` | TLS, X.509, and PSA Crypto | Dual-license path; pin a supported branch and track advisories. |
| `hathach/tinyusb` | USB classes and ISR-to-task deferral patterns | MIT; verify exact USB IP, endpoint memory, cache/DMA, and descriptors. |
| `lwip-tcpip/lwip` | Embedded TCP/IP and pbuf ownership patterns | BSD-style; verify canonical version, core locking, checksums, and port patches. |
| `littlefs-project/littlefs` | Power-loss-aware MCU filesystem | BSD-3-Clause; prove block geometry, wear, migration, and corruption recovery. |
| `nanopb/nanopb` | Bounded Protocol Buffers for C | zlib-style; pin generator/runtime and cap all variable-length fields. |
| `CANopenNode/CANopenNode` | CANopen object dictionary, heartbeat, EMCY, and SDO/PDO patterns | Apache-2.0; verify generated data and timing ownership. |
| `OpenCyphal/libcanard` | Cyphal/CAN transport and bounded queues | MIT; pin an accepted release and DSDL toolchain. |
| `micro-ROS/micro_ros_stm32cubemx_utils` | STM32CubeMX and ROS 2 transport evidence | Apache-2.0 plus notices; upstream readiness does not establish product suitability. |
| `ThrowTheSwitch/CMock` and `meekrosoft/fff` | Generated mocks and header-only fakes | MIT; do not infer ISR/DMA/RTOS correctness from mocks. |
| `pyocd/pyOCD` | Scriptable Cortex-M programming/debug and CMSIS-Pack support | Apache-2.0; pin packs and flash algorithms. |
| `openocd-org/openocd` | SWD/JTAG/GDB and scripted reset/flash/debug | GPL-family tool distribution; verify adapter and target configuration. |
| `renode/renode` | Deterministic virtual MCU/system tests | MIT core with separate libraries; unmodeled timing/analog/electrical behavior is unverified. |
| `Open-CMSIS-Pack/*` | Reproducible pack and csolution/cbuild workflows | Apache-2.0; preserve resolved metadata and license evidence. |
| `apache/nuttx` | POSIX RTOS, STM32 drivers, simulator, and integration evidence | Apache-2.0-compatible; use as architecture evidence unless selected as a dependency. |
| `RIOT-OS/RIOT` | Low-power IoT, STM32 peripheral, and network patterns | LGPL-2.1 core with separate packages; inspect per-file SPDX. |
| `cppcheck-opensource/cppcheck` | Version-pinned static and MISRA-addon analysis | GPL-3.0-or-later tool; confirm findings against compiler/manual evidence. |

See `open-source-ecosystem-gaps.md` for the capability matrix and `supply-chain-and-verification.md` for required integration evidence.

## Source tiers

1. Exact datasheets, reference manuals, errata, schematics, safety manuals, and measured board behavior are normative.
2. Versioned vendor SDKs, evaluation firmware, and application notes are qualified examples.
3. Maintained open source is implementation evidence only after revision, license, target, issue history, and tests are reviewed.
4. Forums, issues, videos, and blogs are leads; trace claims to primary evidence or reproduce them before encoding a rule.

For Modbus, use the current Modbus Organization application and Serial Line specifications and accept their license terms. For a proprietary servo protocol, the vendor manual and captured behavior override generic examples.

## Adoption gate

1. Prefer exact vendor documentation and currently shipped project code.
2. Prefer permissively licensed, maintained upstreams with releases and security history.
3. Pin a release/commit; never depend on an unpinned default branch for production reproduction.
4. Copy the smallest justified pattern and retain required notices.
5. Review bounds, concurrency, timing, endianness, units, and safe state for this board; upstream code is not hardware proof.
6. Add host and target tests that demonstrate why the adopted pattern is correct here.
7. Record the component inventory/SBOM entry, local patches, vulnerability-review date, upgrade owner, and removal or rollback path.
