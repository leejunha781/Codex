# ADC-based motor measurement and calibration

Use this reference for current, DC-bus voltage, back-EMF, potentiometer/command input, temperature, torque proxy, or other analog feedback used to compensate or protect motor operation. ADC self-calibration corrects ADC-internal error; it does not calibrate the shunt, amplifier, divider, sensor, reference, wiring, timing, or plant.

## Measurement contract

For every channel record signal purpose, source impedance, anti-alias filter, protection/clamps, ADC instance/channel, sample time, resolution, alignment, trigger, DMA/injected group, expected range, safe range, scaling units, calibration method, bandwidth, latency, and stale-data behavior.

Use the exact MCU datasheet/reference manual for maximum ADC clock, sampling constraints, VREFINT/temperature calibration addresses and calibration voltage, differential/single-ended rules, and family-specific self-calibration sequence.

## Layered calibration

1. **ADC internal**: run the family-specific self-calibration only in a legal state; respect regulator/startup and single-ended/differential requirements.
2. **Reference/supply**: measure VREFINT and derive VDDA using the device's factory-calibration definition. Do not hard-code a calibration address or voltage across STM32 families.
3. **Channel zero**: with PWM disabled and a defined zero-current condition, average enough samples to estimate each current-channel offset; reject saturation, excessive noise, or disagreement.
4. **Gain/linearity**: use traceable external points or production fixtures to fit gain and, only when justified, piecewise correction. Record fixture uncertainty and temperature.
5. **Runtime compensation**: apply calibrated conversion, plausibility checks, temperature compensation, filtering, and stale-data deadlines before publishing engineering units.
6. **Persistence**: store version, board/MCU identity, timestamp or build context, coefficients, units, temperature range, sample count, quality metrics, and CRC. Validate before use and keep safe defaults.

For an unsigned N-bit ADC, a common starting conversion is `V = raw * VDDA / (2^N - 1)`. For shunt current, a common model is `I = (Vadc - Vzero) / (Rshunt * amplifier_gain)`. Derive polarity, offset, gain, saturation, and units from the actual circuit; use checked arithmetic and state the quantization error.

## Motor-control sampling

- Trigger phase-current conversions from the PWM timer at a quiet, observable point derived from topology, deadtime, blanking, switching edges, and amplifier settling. “Middle of PWM” is not universally valid.
- Use injected/simultaneous ADC modes only when supported by the exact MCU and required by the sensing topology. Document trigger-to-sample latency and channel skew.
- Keep the fast current loop separate from slow Vbus, thermistor, and UI telemetry. Prove DMA/ISR ownership and the age of every sample used by control.
- For one-shunt/three-shunt/isolated-current-sensor topologies, use the matching MCSDK/reference design as an implementation baseline and verify reconstruction windows under overmodulation.
- Back-EMF and sensorless estimation require phase state, blanking, filtering, and speed-dependent observability. Never treat raw ADC threshold crossing as a general-purpose commutation solution.

## Motor identification and compensation

- Perform resistance, inductance, flux-linkage, encoder offset/direction, Hall order, or friction identification only in an explicit calibration state with current-limited power, rotor/workcell control, timeout, emergency stop, and independent overcurrent protection.
- Validate measured parameters against motor/vendor limits and repeatability bounds before committing them.
- Apply temperature/current/voltage compensation gradually and within bounded authority. A calibration routine must not silently raise current or disable a protection threshold.
- Separate calibration quality from control readiness. Failed, stale, wrong-board, or out-of-range calibration must lead to a documented reduced mode or `SAFE_DISABLED`.

## Filtering and diagnostics

- Choose analog and digital filters from required bandwidth and control-loop phase margin. Record group delay; do not add smoothing solely to make plots look stable.
- Detect rail saturation, implausible slew, open/short sensor, channel mismatch, reference failure, noise/RMS excess, missing DMA completion, and stale sequence counters.
- Publish raw value, calibrated value, timestamp/sequence, validity, saturation flags, and calibration version for diagnostics.

## Required tests

- Host golden vectors for scaling, offset, gain, saturation, signedness, overflow, coefficient CRC/version, and fixed-point error.
- Target shorted-input/known-voltage points, VREFINT stability, zero-current drift, multi-point fixture comparison, temperature sweep where required, and injected-trigger timing.
- HIL overcurrent, sensor open/short, Vbus UV/OV, thermistor fault, DMA freeze, noisy reference, wrong calibration blob, brownout during save, and reset during calibration.
- Compare an independent DMM/scope/current probe against firmware values over the operating envelope and retain uncertainty and raw captures.

## Authoritative and open references

- ST AN2834, ADC accuracy optimization: https://www.st.com/resource/en/application_note/an2834-how-to-optimize-the-adc-accuracy-in-the-stm32-mcus-stmicroelectronics.pdf
- STM32 motor-control ecosystem: https://www.st.com/content/st_com/en/ecosystems/stm32-motor-control-ecosystem.html
- X-CUBE-MCSDK: https://www.st.com/en/embedded-software/x-cube-mcsdk.html
- moteus calibration/reference implementation: https://github.com/mjbots/moteus
- VESC firmware sensing and motor-detection implementation: https://github.com/vedderb/bldc
- SimpleFOC current-sense and motor setup examples: https://github.com/simplefoc/Arduino-FOC

Open-source algorithms are implementation evidence, not calibration standards. Check license, board topology, ADC timing, scaling, and protection independently.
