# STM32 Driver Reference

## CubeMX Checklist

- Clock tree matches crystal / PLL constraints in datasheet
- NVIC priorities: safety/fault highest; communication mid; UI/logging lowest
- DMA: request mapping correct; memory/peripheral width; circular vs normal
- After regenerate: re-apply custom code in `USER CODE BEGIN/END` only, or keep drivers outside generated tree

## Minimal Driver Skeleton

```c
/* drv_uart.h */
#pragma once
#include <stdint.h>
#include <stddef.h>

typedef struct {
    void *huart; /* HAL handle or register base — project-specific */
    uint8_t *rx_dma;
    size_t rx_dma_len;
} drv_uart_t;

int drv_uart_init(drv_uart_t *dev);
int drv_uart_start_rx_dma(drv_uart_t *dev);
size_t drv_uart_read(drv_uart_t *dev, uint8_t *out, size_t max);
void drv_uart_irq_handler(drv_uart_t *dev);
void drv_uart_dma_irq_handler(drv_uart_t *dev);
```

App owns lifecycle; driver owns registers/IRQ side effects.

## ISR / DMA Latency Budget

Define before coding high-rate paths:

| Budget item | Target (fill per product) | Measure |
|-------------|---------------------------|---------|
| ISR entry → clear + enqueue | ≤ X µs | GPIO toggle + scope |
| DMA half/full callback | ≤ Y µs | same |
| Control loop period | T | timer capture |
| Worst-case jitter | ≤ Z µs | long capture |

If budget unknown, instrument first — do not guess.

## Motor Control Notes

| Type | Typical interface | Driver must expose |
|------|-------------------|--------------------|
| DC | PWM + DIR + current | duty, direction, fault |
| BLDC | 6-step / FOC + Hall/BEMF | commutation state, current, stall |
| STEP | Pulse/Dir + accel profile | pulse rate, home/limit, missed-step detect |
| SERVO | Position/vel/torque + bus | setpoint, following error, ready/fault |

Do not claim FOC/PID authorship unless implementing it in this task.

## Serial / Fieldbus

- **RS-485:** DE/RE timing vs UART TX complete; termination/bias; timeout
- **RS-232:** ground reference; cable length/noise
- **Ethernet:** PHY link → MAC/DMA → stack → reconnect/watchdog
- Prefer ring buffers + state machine parsers over blocking reads

## HardFault Triage

1. Halt in debugger; read CFSR, HFSR, MMFAR, BFAR
2. Dump stacked frame (R0–R3, R12, LR, PC, xPSR)
3. Map PC to `.map` / disassembly
4. Check stack watermark and concurrent ISR nesting

## Official Docs (start points)

- [STM32 mainstream MCUs](https://www.st.com/en/microcontrollers-microprocessors/stm32-mainstream-mcus.html)
- Part-specific Datasheet + RM + Errata (always required)
- Cube HAL user manual for the series in use
