# Evidence Catalog and Adoption Boundaries

Verify the current canonical release, commit, license, security status, and target support before use. Links identify evidence; they are not automatic approval to copy code.

Cross-check `stm32-device-driver/references/open-source-catalog.md` and `stm32-collaborative-development/references/open-source-qualification.md`. Use one shared qualification record rather than parallel approvals.

## Primary and vendor evidence

- ST X-CUBE-MCSDK and STM32 motor-control ecosystem: official PMSM/FOC tooling and examples. Confirm supported STM32, board, middleware license, generated configuration, and safety limits.
  - https://www.st.com/en/embedded-software/x-cube-mcsdk.html
  - https://www.st.com/content/st_com/en/ecosystems/stm32-motor-control-ecosystem.html
- Arm CMSIS-DSP: Apache-2.0 filtering, statistics, interpolation, and PID primitives. Prove fixed-point scaling, overflow, saturation, state initialization, and timing on the chosen core.
  - https://github.com/ARM-software/CMSIS-DSP
  - https://arm-software.github.io/CMSIS-DSP/main/group__PID.html
- TI ADS1232 data sheet and ADS1232REF guide: bridge ADC behavior, settling/noise, and calibrated weighing practice. Use the exact installed converter data sheet.
  - https://www.ti.com/lit/ds/symlink/ads1232.pdf
  - https://www.ti.com/lit/ug/sbau120b/sbau120b.pdf
- Analog Devices no-OS: vendor-maintained C driver reference; qualify the exact device and license at file level.
  - https://github.com/analogdevicesinc/no-OS

## Community implementation evidence

- Klipper pressure advance and Marlin Linear Advance: useful evidence for transport-pressure compensation and per-material tuning. They target 3D-printer extrusion and do not prove industrial liquid/adhesive dispensing accuracy. Klipper/Marlin licensing and target assumptions require review.
  - https://github.com/Klipper3d/klipper/blob/master/docs/Pressure_Advance.md
  - https://github.com/MarlinFirmware/Marlin
- Hydromisc: open-loop dosing maps and calibration-by-volume/weight examples; agricultural dosing is not a production STM32 reference.
  - https://github.com/hydromisc/hydromisc
- Decent Open Scale and SparkFun OpenScale: weighing architecture and calibration examples; validate ADC, timing, filtering, license, and electrical design independently.
  - https://github.com/decentespresso/openscale
  - https://github.com/sparkfun/OpenScale
- Poseidon, OpenSourceSyringePump, and Wenzel syringe-pump controller: motion and syringe-pump workflow references. Confirm licenses, mechanics, accuracy model, and hardware scope; GPL projects are architectural evidence unless the product license is compatible.
  - https://github.com/pachterlab/poseidon
  - https://github.com/LaubachLab/OpenSourceSyringePump
  - https://github.com/wenzel-lab/syringe-pumps-and-controller
- SimpleFOC, ODrive, VESC, moteus, and Grbl: motor-control or deterministic-motion references. None is drop-in proof for the exact STM32 dispenser, drive, mechanics, or safety case.
  - https://github.com/simplefoc/Arduino-FOC
  - https://github.com/odriverobotics/ODrive
  - https://github.com/vedderb/bldc
  - https://github.com/mjbots/moteus
  - https://github.com/gnea/grbl

## Research evidence

- Gravimetric dosing studies support feed-forward plus iterative/shot-to-shot learning when scale resolution and process delay make fast feedback unsuitable:
  - https://onlinelibrary.wiley.com/doi/10.1155/2018/9425902
  - https://pmc.ncbi.nlm.nih.gov/articles/PMC8510936/
- Adaptive gravimetric fluid dispensing and high-viscosity modeling identify viscosity-dependent PWM, stopping delay, transient delay, and model calibration:
  - https://www.sciencedirect.com/science/article/pii/S1319157811000358
  - https://arxiv.org/abs/2210.10747
- Two-component sources identify startup mixing-ratio transients and temperature/viscosity sensitivity:
  - https://publica.fraunhofer.de/entities/publication/d33a0749-b3d3-48ee-a05c-cd6ff2b83278
  - https://www.mdpi.com/2227-9717/10/5/951

Research papers can justify hypotheses and test designs, not copy-ready production algorithms. Patent documents are background only and require separate legal review before practicing claims.

## Qualification record

For every adopted component record canonical upstream, immutable tag and commit, file-level license and notices, transitive dependencies, security/advisory status, supported target, configuration, local patches, test evidence, SPDX or CycloneDX SBOM entry, upgrade owner, rollback/removal plan, and unresolved hardware claims.
