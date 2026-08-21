# STM32 BSP, Startup, Bootloader, and Boot Images

Use this reference for board-support packages, reset/startup code, ROM or custom bootloaders, application chain-loading, DFU/IAP, and release-image generation. Always select documents for the exact STM32 part number and revision; boot modes, system-memory addresses, flash geometry, cache behavior, TrustZone, and protocol support vary by family and sometimes by suffix.

## Contents

1. Separate the layers
2. Reset and startup contract
3. Choose the boot architecture explicitly
4. Flash and image contract
5. Update state machine and power-loss safety
6. Safe application handoff
7. Deterministic image generation
8. BSP and boot tests
9. Primary and open-source references

## 1. Separate the layers

| Layer | Responsibility | Do not hide here |
|---|---|---|
| CMSIS device/startup | vector table, initial MSP, `Reset_Handler`, `SystemInit`, C runtime entry | board policy and update transport |
| MCU HAL/LL adapter | clocks, GPIO, flash, CRC/hash, watchdog, cache, MPU | application state machines |
| BSP | named board resources and safe initialization order | guessed schematic nets or generic peripheral policy |
| Boot manager | image selection, validation, update state, rollback, recovery, jump | transport-specific packet parsing when separable |
| Transport | UART/I2C/SPI/CAN/FDCAN/USB/storage download | image authenticity policy |
| Image tool | ELF inspection, BIN/HEX conversion, header/TLV, hash/sign, manifest | private keys in the repository or device firmware |

A BSP names real board resources such as `BSP_MOTOR_EN`, `BSP_BOOT_REQUEST`, LEDs, buttons, external flash, and watchdog. It must map them to confirmed schematic nets, declare polarity and reset state, and expose idempotent initialization plus a safe-state function. Keep family HAL and board policy behind interfaces so host tests can replace them.

After configuring GPIO modes and pulls, but before deasserting any motor/power disable or enabling a driver, sample and validate every safety/interlock-classified input. If the fail-safe condition is not met, keep outputs disabled and latch the first reason. Follow `input-subsystems.md` section 9.

## 2. Reset and startup contract

CMSIS startup normally supplies the initial MSP, exception/IRQ vector table, weak default handlers, and a reset handler that calls `SystemInit()` and then the C/C++ runtime entry. Verify the toolchain startup implementation also copies initialized data, zeros BSS, initializes runtime libraries, and eventually calls `main()`.

For every boot image, verify:

- vector table alignment and location permitted by the exact Cortex-M and STM32 reference manual;
- agreement among the boot alias, linker VMA, and vector-table location; when an offset application runs on a core with VTOR, set VTOR before enabling exceptions, and on a core without VTOR place/remap vectors using the exact family-supported mechanism;
- linker/scatter/ICF memory regions, stack/heap symbols, load addresses (LMA), and run addresses (VMA);
- flash sector/page boundaries, write granularity, erased value, ECC constraints, and protected regions;
- FPU, cache, MPU, TrustZone secure/non-secure state, dual-core ownership, and external-memory initialization when applicable;
- watchdog state across reset and handoff;
- reset-cause capture before flags are cleared;
- outputs remain electrically safe from reset through BSP initialization.

Do not replace the vendor CMSIS startup file with a generic vector table. IRQ ordering and names must match the exact device header.

## 3. Choose the boot architecture explicitly

1. **ST ROM/system-memory bootloader:** immutable factory code selected by boot pins, option bytes, or supported software entry. Use AN2606 for the exact device and the protocol note for USART, CAN, USB DFU, I2C, SPI, or FDCAN. Do not assume every interface or pin mapping exists on every part.
2. **STM32 OpenBootloader:** customizable STM32Cube middleware compatible with ST boot protocols. Use when the product needs a modifiable host-facing downloader without inventing a protocol.
3. **Custom IAP/DFU boot manager:** use only with a written flash map, update-state machine, power-loss strategy, image contract, recovery path, and test plan.
4. **Secure boot/update:** prefer a maintained design such as MCUboot, ST MCUboot integration, SBSFU, or TF-M for authenticity, anti-rollback, and recovery. Security depends on hardware protection, immutable roots of trust, and key custody—not only a signature function.

Any boot-request or recovery input must have a confirmed inactive bias, timed qualification, maximum bootloader dwell or authenticated recovery policy, and protection against a stuck input causing an endless boot/update loop.

## 4. Flash and image contract

Define one versioned, endian-explicit descriptor. Typical fields are:

- magic and header format version;
- target product/board and MCU compatibility ID;
- image semantic version and security counter;
- payload offset, size, load address, entry/vector address;
- flags for compression, encryption, secure/non-secure image, and dependencies;
- hash algorithm and payload digest;
- signature algorithm, key ID, and authenticated metadata length;
- update slot and minimum compatible bootloader version.

Validate all arithmetic before flash access:

```c
bool boot_range_valid(uint32_t base, uint32_t size,
                      uint32_t slot_base, uint32_t slot_size,
                      uint32_t program_align)
{
    if ((size == 0u) || (base < slot_base) ||
        (program_align == 0u) ||
        ((base % program_align) != 0u) ||
        ((size % program_align) != 0u)) {
        return false;
    }
    /* Subtraction form avoids an overflowing base + size check. */
    return (size <= slot_size) && ((base - slot_base) <= (slot_size - size));
}
```

CRC detects accidental corruption; it does not establish authenticity. For production secure boot, authenticate the header fields that influence execution as well as the payload, keep production private keys outside source control and build logs, and enforce a monotonic security counter where rollback is prohibited.

If image encryption is enabled, define the exact decrypt engine or software path, key provisioning and custody, plaintext destination, authenticated metadata, option-byte/security configuration, and secure/non-secure placement. Reject the release artifact when any target-specific prerequisite is unverified.

If compression is enabled, define the algorithm/version, maximum expanded size, bounded workspace, authenticated uncompressed metadata, destination-overlap rules, decompression failure handling, and power-loss recovery. Verify the expanded image's digest and ranges before activation.

## 5. Update state machine and power-loss safety

Use explicit states such as `EMPTY`, `DOWNLOADING`, `DOWNLOADED`, `VERIFIED`, `PENDING_TEST`, `CONFIRMED`, and `REJECTED`. Make every flash transition monotonic under the device's 1-to-0 programming rules, and design each erase/program boundary for reset recovery.

- Never erase the only bootable image before a verified recovery path exists.
- Verify bounds before erase, then program in device-supported aligned units.
- Hash/read back the stored image, not only the received buffer.
- A/B or swap designs must resume safely after reset at every sector boundary.
- A trial image must confirm itself only after bounded health checks; otherwise revert.
- Preserve the first failure reason and rate-limit repeated failed updates.
- Protect the bootloader, immutable public key/hash, security counter, and flash map using the exact device mechanisms.

MCUboot image trailers and `imgtool --pad`/slot-size rules are format requirements, not optional padding. Do not hand-create them without following the version-matched MCUboot documentation.

## 6. Safe application handoff

Before a custom jump, validate at minimum that the candidate vector address is aligned and inside the application slot, the initial MSP is 8-byte aligned and matches the application's declared linker/manifest stack region with adequate downward-growth space, the reset vector has the Thumb bit set and points into executable image memory, and image integrity/authenticity has passed. Reject an MSP in heap, unrelated RAM, or a region unavailable to that security/core context.

The family-specific handoff commonly includes:

1. prevent new work and put external outputs in a safe state;
2. stop transport/DMA activity and disable its interrupts;
3. disable SysTick and relevant NVIC interrupts, then clear pending state;
4. apply required cache/barrier/MPU/TrustZone cleanup for the exact core;
5. set the vector-table base when supported and required;
6. set MSP from vector word 0 and branch to vector word 1 as a non-returning call.

Define watchdog continuity explicitly. If an IWDG cannot be stopped after the bootloader starts it, the application must know the running timeout/prescaler and service it early enough during reset initialization; otherwise defer starting it until a documented ownership point. Validate both IWDG and WWDG behavior for cold boot, warm handoff, failed application start, and watchdog-reset recovery.

Sequential bootloader and application execution may legitimately reuse the same RAM. Prove instead that no bootloader stack, DMA, interrupt, callback, or handoff data is still live when application startup overwrites RAM, and that the application linker layout does not overwrite retained mailbox/crash/update state that the contract requires.

For TrustZone parts, validate SAU/IDAU/MPU attribution, secure/non-secure vector tables, NSC veneers, non-secure MSP and reset target, and permitted transition instructions. For dual-core parts, document which core owns validation and update state, how the secondary core is held/reset/released, shared-memory/cache coherency, and each core's vector table. Use a part-specific ST example or application note for the handoff.

Do not copy a jump snippet across STM32 families. Cortex-M0 variants, VTOR availability, TrustZone, dual-core parts, caches, USB, external flash, and active watchdogs change the contract. Test the handoff from cold reset, warm reset, active communications, pending interrupts, and watchdog-reset conditions.

## 7. Deterministic image generation

Treat the linked ELF/AXF as the canonical build artifact. Preserve it with the map file and tool versions. Generate delivery files in a deterministic post-link pipeline:

1. fail the link if boot/application regions overlap or exceed slot budgets;
2. inspect section addresses and the vector table in the ELF;
3. convert to address-bearing Intel HEX/S-record or raw BIN with the toolchain's `objcopy`/`ielftool` equivalent;
4. for BIN, record the required base address separately and reject ambiguity;
5. add the versioned header/TLV and required erased-value padding;
6. hash and sign in a controlled signing environment;
7. re-parse the final file independently and verify size, ranges, header, digest, signature, and security counter;
8. emit a manifest containing product ID, version, source revision, tool versions, file hashes, load address, slot size, and signing-key ID—not a private key;
9. program a clean target and verify readback, boot, trial confirmation, rollback, and power-loss recovery.

GNU `objcopy --gap-fill` fills gaps between sections and `--pad-to` extends to an address. Use them only with a reviewed memory map; accidental padding can produce a huge or destructive raw binary. STM32CubeProgrammer accepts BIN, ELF/AXF/OUT, Intel HEX, and S-record, but a BIN has no embedded target address and must be supplied with the correct address.

## 8. BSP and boot tests

- compile startup and linker variants with warnings and map-file budget checks;
- assert BSP safe-state polarity before enabling clocks or alternate functions;
- test erased, truncated, oversized, wrong-target, wrong-version, bad-hash, bad-signature, and rollback images;
- inject reset during every erase/program/swap phase;
- test boot-request inputs stuck active, bouncing, floating, and asserted during brownout;
- test recovery with unavailable/corrupt external storage and interrupted transport;
- verify watchdog, reset cause, fault log, and boot-attempt counters;
- confirm no motor/power output glitches during reset, update, rollback, or application jump.

## Primary and open-source references

- [Arm CMSIS startup file](https://arm-software.github.io/CMSIS_5/5.8.0/Core/html/startup_c_pg.html)
- [Arm CMSIS use and vector remap](https://arm-software.github.io/CMSIS_5/Core/html/using_pg.html)
- [ST AN2606 system-memory boot mode](https://www.st.com/resource/en/application_note/an2606-introduction-to-system-memory-boot-mode-on-stm32-mcus-stmicroelectronics.pdf)
- [ST OpenBootloader documentation](https://dev.st.com/stm32cube-docs/mw-open-bootloader/7.0.0/en/docs/markup/getting_started.html)
- [ST OpenBootloader source](https://github.com/STMicroelectronics/stm32-mw-openbl)
- [STM32CubeProgrammer](https://www.st.com/content/st_com/en/stm32cubeprogrammer.html)
- [ST AN5056 X-CUBE-SBSFU integration](https://www.st.com/resource/en/application_note/an5056-integration-guide-for-the-xcubesbsfu-stm32cube-expansion-package-stmicroelectronics.pdf)
- [ST MCUboot integration](https://github.com/STMicroelectronics/stm32-mw-mcuboot)
- [MCUboot design](https://docs.mcuboot.com/design.html)
- [MCUboot imgtool](https://docs.mcuboot.com/imgtool.html)
- [Zephyr binary signing](https://docs.zephyrproject.org/latest/build/signing/index.html)
- [GNU objcopy](https://sourceware.org/binutils/docs/binutils/objcopy.html)
