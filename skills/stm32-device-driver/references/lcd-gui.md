# STM32 LCD GUI architecture

Use this reference to build an event-driven STM32 GUI with a desktop-like separation of views, commands, models, and services. “Like C# or MFC” means comparable application structure and iteration flow, not source/API compatibility or desktop memory/process semantics.

## Framework choice

| Choice | Prefer when | Constraints to record |
|---|---|---|
| TouchGFX | STM32-only product, Designer workflow, supported board setup, simulator, tight integration with LTDC/DMA2D/STM32Cube | version/license, supported MCU/toolchain, generated-code boundaries, custom-board abstraction layer |
| LVGL | portable open-source GUI, broad display/touch ecosystem, C API, custom HAL/RTOS integration | pinned major version, MIT license at selected revision, display/input port ownership, memory configuration |
| Thin custom UI | tiny display or fixed screens with strict resource limits | text/font/image pipeline, invalidation, input state, accessibility/localization burden |

Do not mix two GUI frameworks in one render loop unless the design explicitly partitions displays and resources.

## Desktop-to-embedded mapping

| Desktop concept | Embedded mapping |
|---|---|
| Window/form | screen/view with explicit lifetime |
| Event handler | bounded callback that posts a typed command |
| View model/document | portable model/service state, independent of widgets |
| UI thread | single GUI owner task or main-loop context |
| Background worker | RTOS task/ISR that publishes bounded messages; it never touches widgets directly |
| Timer | framework timer or monotonic deadline; no blocking delay in UI code |
| Resource bundle | generated fonts/images/translations in internal or external flash with versioned tooling |

Use Model-View-Presenter, MVVM-like bindings, or a reducer/message architecture. Keep motor control and safety supervision independent of the GUI; the UI requests commands and displays confirmed state. A missing/stalled GUI must not disable protection or keep a motor energized.

## Board bring-up order

1. Prove clocks, MPU/cache, SDRAM/PSRAM, external flash, display reset/power/backlight, and touch-controller electrical interfaces without the GUI framework.
2. Prove solid-color and test-pattern output; verify pixel clock, porches, polarity, color order, orientation, tearing, and frame rate.
3. For SPI/QSPI displays, verify controller ID/init sequence, command/data GPIO, window addressing, maximum bus rate, DMA completion, and partial updates.
4. Size framebuffer(s): `width * height * bytes_per_pixel`, plus stride/alignment and optional double buffering. Prove the selected memory is visible to CPU, DMA2D, LTDC/DSI, and DMA.
5. Add the framework abstraction layer, then input devices, fonts/images, animations, localization, and product services.

## Rendering and concurrency

- Select direct, partial, or full rendering from RAM/bandwidth measurements. Do not assume double buffering fits or eliminates tearing.
- Maintain D-cache coherency for framebuffer and DMA regions on applicable MCUs; align and isolate cache-maintained buffers.
- Synchronize buffer swap/flush completion to LTDC/DSI/SPI events. Never signal “flush ready” before the transfer completes.
- Give the GUI one owner. Publish sensor/motor updates through a queue, mailbox, or immutable snapshot and coalesce high-rate telemetry.
- Bound frame time, dirty area, queue depth, string formatting, image decode, and external-flash latency. Avoid heap allocation in repeated paint/update paths.
- Treat touch as an input device with calibration matrix, debounce/filtering, rotation mapping, press/release semantics, and fault behavior.

## Source layout

```text
App/             product model, typed commands, navigation, presentation
Gui/             generated views/widgets and handwritten presenters/controllers
Display/         panel init, framebuffer, LTDC/DSI/SPI flush, DMA2D, backlight
Input/           touch/button/encoder adapters and calibration
Platform/        RTOS, time, memory/cache, external flash
Diagnostics/     frame time, missed-vsync, flush timeout, queue high-water mark
Tests/           host model/presenter tests and target visual/performance tests
```

## Verification

- Host/simulator: navigation, presenter/model logic, localization bounds, invalid values, command confirmation, and error screens.
- Target: boot-to-first-frame, color bars, touch grid, orientation, clipping, fonts, animation worst case, memory high-water mark, frame/flush deadlines, and 24-hour soak.
- Faults: display/touch absent, DMA timeout, external-memory error, corrupted asset, queue overflow, motor communication loss, brownout, and watchdog reset.
- Capture screenshots plus logic-analyzer/trace timing. Visual correctness alone does not prove bus timing or cache coherency.

## Authoritative and community entry points

- TouchGFX development and board bring-up: https://support.touchgfx.com/docs/development/development-introduction and https://support.touchgfx.com/docs/development/board-bring-up/board-introduction
- TouchGFX abstraction layer: https://support.touchgfx.com/docs/development/touchgfx-hal-development/touchgfx-al-development-introduction
- TouchGFX simulator/debugging: https://support.touchgfx.com/docs/development/ui-development/working-with-touchgfx/simulator
- LVGL STM32 integration: https://docs.lvgl.io/master/integration/chip_vendors/stm32/index.html
- LVGL SPI display guide: https://docs.lvgl.io/master/integration/chip_vendors/stm32/lcd_stm32_guide.html
- LVGL source: https://github.com/lvgl/lvgl

Pin documentation to the selected framework release; APIs and generated layouts change between major versions.
