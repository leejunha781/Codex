# STM32 bus and register-device driver patterns

Use this reference for external ICs reached through I2C, SPI, UART/RS-232/RS-485, CAN/FDCAN, or GPIO/PWM. Treat the exact MCU reference manual, errata, board schematic, external-device datasheet, and electrical measurements as the contract. STM32Cube examples establish API usage, not the target board's wiring or safety.

## Evidence and design order

1. Record exact MCU and external-IC order codes, document revisions, bus instance, pins/alternate functions, voltage domains, pull-ups, isolation, termination, chip-select/address straps, reset/enable/fault polarity, and clock limits.
2. Extract a register table containing address, reset value, access type, reserved-bit policy, field encoding, side effects, clear semantics, and legal state. Do not encode registers from memory.
3. Define transport semantics: blocking or asynchronous, timeout, retry eligibility, bus recovery, DMA/cache requirements, ISR/task ownership, transaction serialization, and cancellation.
4. Keep wire encoding and state logic portable. Put HAL/LL calls, GPIO chip select, delay, cache maintenance, and RTOS primitives behind adapters.
5. Read back identity/configuration where supported. Compare masked values, preserve reserved bits when required, and record every mismatch.
6. Enter the safe state after partial initialization, transport loss, impossible status, or repeated timeout.

## Portable interface pattern

```c
typedef enum {
    DEV_OK = 0,
    DEV_E_ARG,
    DEV_E_TIMEOUT,
    DEV_E_BUS,
    DEV_E_ID,
    DEV_E_CONFIG,
    DEV_E_FAULT,
    DEV_E_STATE
} dev_status_t;

typedef dev_status_t (*dev_txrx_fn)(void *ctx,
                                    const uint8_t *tx, size_t tx_len,
                                    uint8_t *rx, size_t rx_len,
                                    uint32_t timeout_ms);

typedef struct {
    dev_txrx_fn txrx;
    void *ctx;
    uint32_t timeout_ms;
} dev_bus_t;
```

Return a typed status; never collapse timeout, NACK, arbitration loss, CRC failure, device fault, and invalid state into one boolean. Add asynchronous start/cancel/completion operations only when the ownership model is documented and tested.

## Register codec rules

- Use fixed-width unsigned types and explicit masks/shifts. Validate a value before shifting it.
- Serialize byte order explicitly; do not transmit packed C structs or bitfields.
- For SPI, define CPOL, CPHA, bit order, word length, maximum frequency, chip-select setup/hold/high time, daisy-chain framing, dummy cycles, and read/write command bits.
- For I2C, define 7-bit address, address straps, repeated-start/stop rules, register-address width and byte order, clock stretching, transfer limits, and stuck-bus recovery. Store the logical 7-bit address in portable code and shift only in an adapter if the selected HAL API requires it.
- For UART/RS-485, define electrical layer separately from framing. Specify DE/RE timing, turnaround, idle detection, inter-byte/frame timeout, CRC, replay/duplicate handling, and collision/backoff policy.
- For CAN/FDCAN, define identifier type, DLC, byte order, signal scaling, freshness, heartbeat, bus-off recovery, acceptance filters, and command authorization. Use a versioned application protocol; a CAN frame is not itself a safe motor command.

## DMA, cache, and concurrency

- Keep DMA buffers alive, aligned, in DMA-accessible memory, and owned for the entire transfer.
- On cache-equipped STM32 parts, derive clean/invalidate operations and memory placement from the exact core/MCU documentation. Align operations to cache lines and prevent unrelated data from sharing maintained lines.
- Serialize a shared bus at transaction level, including chip-select assertion through completion. Never hold a bus lock while waiting on an unrelated task or user callback.
- Keep callbacks bounded. Post an event or copy a small result; parse and update product state outside ISR context.
- Use elapsed-time subtraction for wrap-safe deadlines. Bound retries and do not retry writes with non-idempotent side effects unless the protocol supplies a safe confirmation mechanism.

## Required tests

- Golden-vector encode/decode and register-mask tests on the host.
- Wrong ID, reserved-bit contamination, reset-value mismatch, timeout, NACK, CRC/parity, truncated reply, busy/stuck status, and retry exhaustion.
- Logic-analyzer verification of frequency, mode, byte order, chip select, repeated start, turnaround, inter-frame timing, and error recovery.
- DMA wrap, cache coherency, back-to-back transfers, concurrent callers, cancellation, reset during transfer, and partial-initialization cleanup.

## Normative starting points

- STM32Cube MCU packages and examples: https://github.com/STMicroelectronics/STM32Cube_MCU_Overall_Offer
- STM32 HAL I2C documentation example: https://dev.st.com/stm32cube-docs/stm32c5xx-hal-drivers/2.0.0/en/docs/drivers/hal_drivers/i2c/hal_i2c_overview.html
- STM32Cube part-driver architecture: https://dev.st.com/stm32cube-docs/embedded-software/2.0.0/en/architecture/part-drivers-architecture.html
- ST platform-independent C driver pattern: https://github.com/STMicroelectronics/STMems_Standard_C_drivers

Use the matching family/package release rather than assuming the linked family or HAL generation applies unchanged.
