# STM32 Input Subsystems

Use this reference for digital inputs, EXTI, buttons, switches, key matrices, rotary/quadrature encoders, timer input capture, pulse/frequency measurement, ADC/DMA inputs, serial receive paths, low-power wake inputs, and safety/interlock signals.

## Contents

1. Input contract
2. Select the acquisition method
3. GPIO and EXTI
4. Debounce and gesture state
5. Matrix keypad
6. Encoders and timer inputs
7. ADC and DMA inputs
8. Serial and protocol inputs
9. Safety and boot-selection inputs
10. Verification matrix
11. Primary and open-source references

## 1. Input contract

For every input, record:

- exact MCU pin, package, alternate function, board net, connector pin, voltage domain, and source device;
- active polarity, pull-up/down ownership, open-drain/open-collector behavior, Schmitt characteristics if documented, and external filtering/protection;
- valid, invalid, floating, disconnected, and power-off behavior;
- sampling method, timestamp domain, minimum pulse width, maximum rate, latency, jitter, and debounce/filter time;
- ISR/DMA/task ownership, queue depth, overflow policy, and stale-data timeout;
- safety classification, diagnostic coverage, startup state, and recovery policy.

Never infer polarity or an internal pull from a signal name. Check absolute maximum ratings and digital thresholds for the actual VDD and I/O type; not every STM32 pin is 5 V tolerant, and tolerance may change in analog mode or when unpowered.

## 2. Select the acquisition method

| Input | Preferred mechanism | Key risk |
|---|---|---|
| slow level/status | periodic GPIO sample | floating/stale state |
| human button/switch | GPIO plus timed debounce | bounce and long-press ambiguity |
| asynchronous edge | EXTI ISR to bounded event queue | shared lines, storms, lost edges |
| matrix keypad | timed row scan and ghost policy | ghosting, shoot-through, EMI |
| quadrature encoder | timer encoder mode when available | rollover, illegal transitions, filter delay |
| pulse width/frequency | timer input capture, optional DMA | counter wrap, overcapture, clock error |
| analog sensor | timer-triggered ADC plus DMA | source impedance, sample time, freshness |
| high-rate serial input | DMA/ring buffer plus idle/timeout | overwrite, framing, cache coherency |
| wake/safety input | dedicated EXTI/wakeup/break path | polarity, latch, reset-domain behavior |

Do not perform parsing, logging, blocking HAL calls, long debounce delays, or application actions inside an ISR. Capture the minimum state and timestamp, clear the exact source correctly, and defer policy to a task or main-loop state machine.

## 3. GPIO and EXTI

- Configure mode, pull, speed, and alternate function from the exact datasheet/reference manual and schematic.
- Configure in the safe order required by the family: GPIO input mode and pull, SYSCFG/EXTICR port-to-line mapping, EXTI trigger/mask, clear existing pending state, configure priority, then enable NVIC. Verify the selected port/line with a controlled edge before relying on it for safety.
- Read the input data register for a physical level; do not confuse it with the output latch.
- Clear EXTI pending flags using the family-defined semantics; writing a guessed read-modify-write sequence can clear unrelated shared flags.
- Account for EXTI line sharing and NVIC priority. A pin number can map to only one GPIO port source for a given EXTI line on many families.
- For ISR-to-task communication, use a fixed-size queue or atomic flag/counter with an explicit overflow policy.
- A wake source may behave differently across Stop/Standby modes; verify retention, pull configuration, wake flags, and post-wake reinitialization.

## 4. Debounce and gesture state

Debounce is a time-qualified state transition, not a blocking delay. Use wrap-safe elapsed-time arithmetic and define press, release, long-press, repeat, and chord semantics separately.

```c
typedef struct {
    bool stable;
    bool candidate;
    uint32_t candidate_since_ms;
} input_debounce_t;

bool input_debounce_update(input_debounce_t *s, bool raw,
                           uint32_t now_ms, uint32_t qualify_ms)
{
    if (raw != s->candidate) {
        s->candidate = raw;
        s->candidate_since_ms = now_ms;
    }
    if ((s->stable != s->candidate) &&
        ((uint32_t)(now_ms - s->candidate_since_ms) >= qualify_ms)) {
        s->stable = s->candidate;
        return true;
    }
    return false;
}
```

Initialize `stable` and `candidate` from a real sampled level after GPIO configuration. Test tick wraparound, asymmetric press/release timing when required, rapid chatter, stuck inputs, and reset while pressed.

## 5. Matrix keypad

- Drive at most the intended row state; avoid opposing outputs through two pressed keys.
- Allow electrical settling before sampling columns.
- Use external/internal pulls with confirmed current and sleep behavior.
- Define whether multi-key rollover is supported. Without isolation diodes, detect and reject ambiguous ghost combinations rather than inventing keys.
- Debounce each logical key after matrix reconstruction, not merely each raw column.
- Bound scan frequency, CPU load, event queue, repeat generation, and wake strategy.

## 6. Encoders and timer inputs

For quadrature encoders, prefer timer encoder mode when the chosen STM32 timer/pins support it. Configure input polarity and digital filtering from measured edge rate and noise; excessive filtering loses legitimate transitions. Extend narrow hardware counters with an overflow-safe signed accumulator and sample faster than the maximum ambiguous movement.

For input capture:

- derive the timer clock from the actual clock tree, including APB timer multipliers;
- define prescaler, counter width, capture edge, input filter, and overflow extension;
- check capture flags and overcapture before accepting a sample;
- use modulo subtraction for a single-wrap interval only when the maximum interval proves that assumption;
- reject zero/out-of-range periods before division;
- convert ticks with checked 64-bit intermediates and documented rounding;
- distinguish timeout/no-signal from a valid zero-frequency result.

Safety encoders need plausibility checks, direction agreement, freshness limits, independent overspeed protection, and a defined response to illegal quadrature transitions or cable faults.

## 7. ADC and DMA inputs

Use `adc-motor-calibration.md` for motor feedback. In all ADC paths, prove source impedance versus sample time, channel order, reference/VDDA handling, calibration state, trigger phase, DMA buffer ownership, cache maintenance, saturation, plausibility, and freshness. A DMA completion interrupt means memory transfer completed; it does not prove the sensor or reference is valid.

## 8. Serial and protocol inputs

For UART/RS-485/SPI/I2C/CAN receive paths, use `bus-device-driver-patterns.md`. Treat received length and metadata as untrusted. Frame into bounded buffers, detect overwrite/overrun, validate length/version/CRC/authentication before executing commands, and define timeout and resynchronization. For UART DMA with idle detection, snapshot producer position atomically and handle wrap, cache, and partial-frame semantics.

## 9. Safety and boot-selection inputs

- Sample safety inputs into a fail-safe state before enabling motor or power outputs.
- Prefer hardware break/emergency paths for response times software cannot guarantee.
- Latch first faults; never automatically re-enable solely because an input returned inactive.
- A boot-mode or factory-test input must be debounced or time-qualified, have a defined inactive bias, and be protected from accidental activation.
- Do not let a stuck boot-request input cause an unbounded boot loop; provide a timeout or authenticated recovery policy.

## 10. Verification matrix

- nominal active/inactive levels at minimum/maximum VDD and temperature;
- floating, open, short-to-ground, short-to-supply, reversed connector, and powered-source/unpowered-MCU cases;
- bounce/chatter, minimum pulse, maximum rate, EMI burst, simultaneous inputs, and queue overflow;
- EXTI pending on enable, shared-line activity, priority inversion, interrupt storm, and tick wrap;
- encoder forward/reverse, rollover, illegal transitions, overspeed, cable disconnect, and index alignment;
- input capture overflow/overcapture, no-signal timeout, clock tolerance, and independent instrument correlation;
- ADC DMA freeze, stale samples, saturation, wrong channel order, reference drift, and brownout;
- Stop/Standby wake polarity, retained pulls, wake flags, and post-wake state;
- reset while asserted and safe behavior before configuration completes.

## Primary and open-source references

- [STM32CubeMX initialization code generation manual](https://www.st.com/resource/en/user_manual/dm00104712-stm32cubemx-for-stm32-configuration-and-initialization-c-code-generation-stmicroelectronics.pdf)
- [STM32CubeG4 official GPIO/EXTI examples index](https://github.com/STMicroelectronics/STM32CubeG4/blob/master/Projects/STM32CubeProjectsList.html)
- [ST STM32Cube MCU open-source offer](https://github.com/STMicroelectronics/STM32Cube_MCU_Overall_Offer)
- [Zephyr GPIO keys binding](https://docs.zephyrproject.org/latest/build/dts/api/bindings/input/gpio-keys.html)
- [Zephyr STM32 quadrature decoder binding](https://docs.zephyrproject.org/latest/build/dts/api/bindings/sensor/st%2Cstm32-qdec.html)
- [LVGL input-device overview](https://docs.lvgl.io/master/details/main-modules/indev.html)
