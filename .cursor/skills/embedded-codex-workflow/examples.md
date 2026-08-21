# Embedded ↔ Codex Examples

## Example 1 — New UART DMA driver

**User:** STM32G4 보드에서 RS-485용 UART DMA 수신 드라이버 작성

**Agent actions:**

1. Load `embedded-codex-workflow` + `stm32-device-drivers`
2. Confirm part, USART instance, DE GPIO, DMA channel
3. Implement `drv_uart_rs485_*` with TX-complete → DE drop
4. Provide CubeMX NVIC/DMA notes + scope test (DE vs TX)

## Example 2 — Dispense state machine

**User:** PROCON-style Time mode + suck-back 상태머신 초안

**Agent actions:**

1. Load `taeha-dispense-controller` + `embedded-codex-workflow`
2. Implement `app_sm` with guards and FAULT latch
3. Stub motor/sensor services behind interfaces
4. Unit-test transitions; document public-manual assumptions

## Example 3 — Codex handoff after Claude design

**Claude output:** timing budget table + SM diagram  
**Codex/Cursor:** implement drivers per budget; return measurement hooks (GPIO markers)

Do not re-litigate architecture unless measurements break the budget.
