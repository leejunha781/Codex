# Reviewed STM32 open-source ecosystem gaps

Use this map to find evidence, not to import every project. Review date: 2026-08-18. Recheck the canonical repository, current release, security status, file-level license, notices, and target support at adoption time.

## Coverage map

| Capability gap | Qualified upstream evidence | Default use | Required boundary |
|---|---|---|---|
| Secure boot and recovery | `mcu-tools/mcuboot` 2.4.0, ST SBSFU/MCUboot ports | Image/TLV, signing, swap/revert, simulator patterns | Prove STM32 flash geometry, option bytes, key custody, monotonic rollback, watchdog and reset-at-every-write recovery. |
| Trusted execution and PSA | `TrustedFirmware-M/trusted-firmware-m` | Secure partition, attestation and PSA service architecture | Select an exact TrustZone-capable MCU and vendor platform; review all external directories and secure/non-secure veneers. |
| Cryptography/TLS | `Mbed-TLS/mbedtls` supported LTS/current branch | PSA Crypto, TLS/X.509 implementation | Choose the license path; configure entropy, time, certificate validation, side-channel/fault controls and security-update cadence. |
| USB | `hathach/tinyusb` | Descriptor/class and bounded deferred-event patterns | Exact STM32 USB IP, endpoint RAM, cache/DMA, malformed setup packets, disconnect/reset and class-specific fuzzing. |
| TCP/IP | canonical lwIP on Savannah; `lwip-tcpip/lwip` mirror | pbuf/core-locking/network-service evidence | Pin the ST/Cube port delta, thread model, timeouts, checksum offload, DMA descriptors, cache/MPU and packet fuzzing. |
| Power-loss storage | `littlefs-project/littlefs` 2.11.3 | Block-device contract and atomic update evidence | Exact erase/program geometry, bad-block assumptions, wear budget, sync boundary, migration and corruption recovery. |
| Bounded serialization | `nanopb/nanopb` | Small Protocol Buffers runtime/generator | Pin runtime and generator together; cap all fields, avoid implicit heap use, test unknown/truncated/oversized inputs. |
| Industrial CAN | `CANopenNode/CANopenNode`, `OpenCyphal/libcanard` | CANopen/Cyphal state and transport evidence | Protocol choice is a product decision; pin generated dictionaries/DSDL, define bus-off, queue, transfer-ID, redundancy and recovery. |
| Robotics middleware | `micro-ROS/micro_ros_stm32cubemx_utils` | CubeMX transport/build reference | Upstream says not production-ready; use only after memory, timing, agent-loss, QoS, lifecycle and safety analysis. |
| RTOS cross-check | FreeRTOS Kernel, Zephyr, Apache NuttX, RIOT | Compare scheduler, driver, power and test patterns | Do not mix APIs or copy ports blindly; selected RTOS manual and exact port remain normative. |
| Host doubles | Unity, CMock, FFF | Portable module tests and HAL seams | Mocks prove calls and state logic only; they do not prove interrupt priority, DMA, cache, bus timing or electrical behavior. |
| Virtual target | Renode 1.16.1 | Deterministic CPU/peripheral scenarios, multi-node tests | Record machine model coverage; analog, electrical, cycle timing, cache/DMA and undocumented silicon behavior require target tests. |
| Probe/debug | pyOCD 0.44.1, OpenOCD 0.12.0 | Scriptable flash/reset/GDB flows | Pin flash algorithms/configs, identify probe serial and target, validate connect-under-reset and recovery before automation. |
| Reproducible build | Open-CMSIS-Pack spec/toolbox | Frozen pack resolution and build metadata | Commit resolved pack metadata, archive licenses, pin compiler and generator versions, hash artifacts and forbid silent pack updates. |
| Static analysis | clang-tidy/Clang, Cppcheck 2.20.0 | Complementary diagnostics and MISRA support | Pin suppressions/config, review false positives, correlate with the target compiler, and never treat zero findings as proof. |

## Projects deliberately kept conditional

- `vedderb/bldc`, RIOT and libopencm3 have copyleft obligations; use as evidence unless a documented product license decision permits integration.
- `micro-ROS` states that its STM32CubeMX utility is not production-ready; do not promote a demo integration into a safety-relevant controller without a lifecycle decision.
- `libcanard` current main may describe alpha APIs. Pin a released revision and its DSDL generator rather than following the default branch.
- GitHub mirrors for TF-M, lwIP and OpenOCD are not automatically the canonical release authority. Record the upstream release location.
- Vendor Cube packages contain mixed middleware licenses. A top-level repository license never substitutes for component-level review.

## Evidence captured in the adoption record

Record repository URL, canonical upstream, retrieval date, tag, full commit, selected files, component and file-level licenses, notices, target/toolchain, compile-time configuration, transitive components, local patches, security-review date, adopted pattern, rejected alternatives, test evidence, upgrade owner, and removal/rollback path.

Never describe this catalog as exhaustive. It is a maintained qualification shortlist; a new project can be considered only after the same evidence gate.
