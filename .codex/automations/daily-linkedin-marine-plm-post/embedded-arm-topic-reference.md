# Embedded Hardware & Software Topic Reference — LinkedIn Posts

**Purpose:** Rotate ARM Cortex-M / Cortex-A embedded design content into the Daily LinkedIn pipeline alongside Marine PLM, SATCOM, and freelance topics.

**Rotation rule:** at least **1 in 4** posts uses an embedded hardware/software angle. Prefer connecting embedded work to marine/defense/satellite context when natural (e.g. terminal baseband MCU, shipboard control board, RF modem companion MCU).

---

## Domain pillars (use these angles)

### 1. ARM Cortex-Mx — MCU hardware & firmware
- Board bring-up, clock tree, power domains, reset sequencing
- Bare-metal / RTOS firmware architecture (FreeRTOS, Zephyr, CMSIS)
- Peripheral drivers: GPIO, UART, SPI, I2C, CAN, ADC, DMA, timers, PWM
- Interrupt priority / latency, ISR vs deferred work
- Flash/EEPROM layout, bootloader, OTA update gates
- Low-power modes, watchdog, brown-out detection

### 2. ARM Cortex-Ax — application-class embedded Linux / BSP
- SoC selection, DDR/eMMC, boot chain (SPL → U-Boot → kernel)
- Device tree, pinmux, regulator, clock bindings
- Linux driver model: platform / I2C / SPI / character / network drivers
- Userspace HAL vs kernel driver boundary
- Yocto/Buildroot image reproducibility and SBOM for field updates

### 3. Peripheral device driver coding
- Register maps, HAL layering, DMA + interrupt co-design
- Protocol stacks on MCU (Modbus, NMEA, CANOpen, proprietary RF framing)
- Validation: loopback, logic analyzer evidence, timing budgets
- Defensive coding: timeout, CRC, retry, fault injection tests

### 4. Marine / defense / satellite bridge (preferred when possible)
- Shipboard control board commissioning evidence
- SATCOM terminal companion MCU / modem lock validation
- Naval EMC/EMI constraints on PCB and cabling
- Field RCA: link loss vs MCU hang vs driver race condition

---

## Suggested topic slugs (rotate; do not repeat within 14 days)

| topic-slug | angle |
|------------|-------|
| cortex-m-peripheral-driver-gates | Driver readiness gates before board release |
| cortex-m-bringup-evidence-chain | Bring-up checklist as auditable evidence |
| cortex-a-bsp-device-tree-ownership | Device-tree ownership and field update risk |
| embedded-dma-isr-latency-budget | ISR/DMA latency budgets for real-time I/O |
| embedded-bootloader-ota-acceptance | Bootloader/OTA acceptance before fleet deploy |
| cortex-m-can-uart-driver-validation | CAN/UART driver validation with loopback proof |
| embedded-linux-driver-boundary | Kernel vs userspace HAL for shipboard SoC |
| satcom-companion-mcu-lock-proof | Companion MCU role in modem lock evidence |
| embedded-emc-pcb-commissioning | EMC-aware PCB/layout gates for naval install |
| rtos-watchdog-fault-isolation | Watchdog + fault isolation for field RCA |

---

## Post writing rules (embedded angle)

- English, consultative, insight-led — not chip vendor marketing
- Include at least one concrete execution context: register map ownership, bring-up gate, driver validation evidence, boot/OTA acceptance, latency budget, or marine/SATCOM integration
- Prefer diagrams: peripheral bus map, boot chain, ISR/DMA pipeline, driver layering
- Soft CTA rare; most posts are engineering insight
- Image style: same commissioning-gates / solution-overview aesthetic when using the Figma/reference pipeline (deep navy, 5–7 gates, central hardware scene)

---

## Example hooks

- "A Cortex-M driver that 'works on the bench' still fails sea trial without ISR latency and DMA ownership evidence."
- "Cortex-A bring-up is not finished at U-Boot prompt — device-tree and regulator bindings are the release gate."
- "Peripheral driver coding is acceptance engineering: timeouts, CRC, and loopback proof belong in the closure pack."
