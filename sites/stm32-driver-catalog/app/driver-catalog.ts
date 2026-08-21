export type DriverSeed = {
  name: string;
  slug: string;
  category: string;
  summaryEn: string;
  summaryKo: string;
  applicationEn: string;
  applicationKo: string;
  signature: string;
  declaration: string;
  functionName: string;
  body: string[];
  guideUrl?: string;
  specs?: [string, string][];
  tags?: string[];
  filename?: string;
  headerFilename?: string;
  source?: string;
  header?: string;
};

const wiki = "https://wiki.st.com/stm32mcu/wiki/";

export const halReferenceUrl = "https://www.st.com/resource/en/user_manual/dm00105879-stm32cube-hal-and-lowlayer-drivers-api-reference-manual-stmicroelectronics.pdf";

export const additionalDrivers: DriverSeed[] = [
  { name: "TouchGFX Display HAL", slug: "touchgfx_display_hal", category: "Graphics", summaryEn: "Connect TouchGFX partial-framebuffer blocks to a board display transport.", summaryKo: "TouchGFX 부분 프레임버퍼 블록을 보드 디스플레이 전송부에 연결합니다.", applicationEn: "Custom SPI, FMC and DSI-command display drivers with DMA completion handling.", applicationKo: "DMA 완료 처리가 필요한 사용자 정의 SPI, FMC 및 DSI 명령형 디스플레이 드라이버.", signature: "void", declaration: "void TouchGFX_DisplayHAL_Init(void);", functionName: "TouchGFX_DisplayHAL_Init", body: [], filename: "TouchGFXDisplayDriver.cpp", headerFilename: "TouchGFXDisplayDriver.hpp", source: `// TouchGFXDisplayDriver.cpp — board transport adapter
#include "TouchGFXDisplayDriver.hpp"

static volatile bool transferActive = false;

extern "C" int touchgfxDisplayDriverTransmitActive(void) {
  return transferActive ? 1 : 0;
}

extern "C" int touchgfxDisplayDriverShouldTransferBlock(uint16_t bottom) {
  (void)bottom;
  return transferActive ? 0 : 1;
}

extern "C" void touchgfxDisplayDriverTransmitBlock(
  const uint8_t* pixels, uint16_t x, uint16_t y,
  uint16_t width, uint16_t height
) {
  if (!pixels || width == 0U || height == 0U || transferActive) return;
  transferActive = true;
  BoardDisplay_SetWindow(x, y, width, height);
  if (BoardDisplay_TransmitDMA(pixels, (uint32_t)width * height * 2U) != HAL_OK) {
    transferActive = false;
  }
}

extern "C" void TouchGFX_DisplayTransferComplete(void) {
  transferActive = false;
}`, header: `#ifndef TOUCHGFX_DISPLAY_DRIVER_HPP
#define TOUCHGFX_DISPLAY_DRIVER_HPP
#include "main.h"
#include <stdbool.h>
#ifdef __cplusplus
extern "C" {
#endif
int touchgfxDisplayDriverTransmitActive(void);
int touchgfxDisplayDriverShouldTransferBlock(uint16_t bottom);
void touchgfxDisplayDriverTransmitBlock(const uint8_t* pixels, uint16_t x, uint16_t y, uint16_t width, uint16_t height);
void TouchGFX_DisplayTransferComplete(void);
void BoardDisplay_SetWindow(uint16_t x, uint16_t y, uint16_t width, uint16_t height);
HAL_StatusTypeDef BoardDisplay_TransmitDMA(const uint8_t* pixels, uint32_t bytes);
#ifdef __cplusplus
}
#endif
#endif`, specs: [["LAYER", "TouchGFX HAL"], ["STRATEGY", "Partial framebuffer"], ["TRANSFER", "DMA / board transport"]], tags: ["TOUCHGFX", "DISPLAY HAL", "PARTIAL FB"], guideUrl: "https://support.touchgfx.com/docs/development/touchgfx-hal-development/touchgfx-architecture" },
  { name: "TouchGFX Touch Driver", slug: "touchgfx_touch_driver", category: "Graphics", summaryEn: "Adapt a board touch-controller driver to the TouchGFX sampleTouch interface.", summaryKo: "보드 터치 컨트롤러 드라이버를 TouchGFX sampleTouch 인터페이스에 연결합니다.", applicationEn: "I2C or SPI capacitive and resistive touch controllers with board-specific coordinate reads.", applicationKo: "보드 전용 좌표 읽기를 사용하는 I²C·SPI 정전식 및 저항식 터치 컨트롤러.", signature: "void", declaration: "void TouchGFX_TouchDriver_Init(void);", functionName: "TouchGFX_TouchDriver_Init", body: [], filename: "STM32TouchController.cpp", headerFilename: "STM32TouchController.hpp", source: `// STM32TouchController.cpp — TouchGFX input adapter
#include "STM32TouchController.hpp"

void STM32TouchController::init() {
  BoardTouch_Init();
}

bool STM32TouchController::sampleTouch(int32_t& x, int32_t& y) {
  int32_t rawX = 0;
  int32_t rawY = 0;
  if (!BoardTouch_Read(&rawX, &rawY)) return false;
  if (rawX < 0 || rawY < 0 || rawX >= DISPLAY_WIDTH || rawY >= DISPLAY_HEIGHT) return false;
  x = rawX;
  y = rawY;
  return true;
}`, header: `#ifndef STM32_TOUCH_CONTROLLER_HPP
#define STM32_TOUCH_CONTROLLER_HPP
#include <platform/driver/touch/TouchController.hpp>
#include <stdint.h>

#ifndef DISPLAY_WIDTH
#define DISPLAY_WIDTH 480
#endif
#ifndef DISPLAY_HEIGHT
#define DISPLAY_HEIGHT 272
#endif

extern "C" void BoardTouch_Init(void);
extern "C" bool BoardTouch_Read(int32_t* x, int32_t* y);

class STM32TouchController : public touchgfx::TouchController {
public:
  void init() override;
  bool sampleTouch(int32_t& x, int32_t& y) override;
};
#endif`, specs: [["LAYER", "TouchGFX input"], ["API", "sampleTouch"], ["TRANSPORT", "I2C / SPI board driver"]], tags: ["TOUCHGFX", "TOUCH INPUT", "BOARD ADAPTER"], guideUrl: "https://support.touchgfx.com/docs/development/board-bring-up/how-to/09-touch-controller" },
  { name: "RGB-TFT LCD", slug: "rgb_tft_lcd", category: "Display", summaryEn: "Stream RGB pixel data and timing directly through the LTDC peripheral.", summaryKo: "LTDC 주변장치로 RGB 픽셀 데이터와 타이밍을 직접 출력합니다.", applicationEn: "Medium and large RGB565/RGB888 panels without an integrated framebuffer.", applicationKo: "내장 프레임버퍼가 없는 중대형 RGB565/RGB888 패널.", signature: "LTDC_HandleTypeDef *hltdc, uint32_t framebuffer", declaration: "HAL_StatusTypeDef RGB_TFT_Start(LTDC_HandleTypeDef *hltdc, uint32_t framebuffer);", functionName: "RGB_TFT_Start", body: ["  if (!hltdc || framebuffer == 0U) return HAL_ERROR;", "  if (HAL_LTDC_SetAddress(hltdc, framebuffer, 0U) != HAL_OK) return HAL_ERROR;", "  __HAL_LTDC_ENABLE(hltdc);", "  return HAL_OK;"], specs: [["INTERFACE", "RGB-TFT / LTDC"], ["PIXEL", "RGB565 / RGB888"], ["FRAMEBUFFER", "Internal or SDRAM"]], tags: ["LTDC", "PARALLEL RGB", "DIRECT PANEL"], guideUrl: "https://www.st.com/resource/en/application_note/an4861-introduction-to-lcdtft-display-controller-ltdc-on-stm32-mcus-stmicroelectronics.pdf" },
  { name: "MIPI DSI LCD", slug: "mipi_dsi_lcd", category: "Display", summaryEn: "Start a one- or two-lane MIPI-DSI link in command or video mode.", summaryKo: "1·2레인 MIPI-DSI 링크를 명령 또는 비디오 모드로 시작합니다.", applicationEn: "Low-pin-count mobile-style panels using DSI video or command mode.", applicationKo: "DSI 비디오 또는 명령 모드를 사용하는 저핀수 모바일형 패널.", signature: "DSI_HandleTypeDef *hdsi", declaration: "HAL_StatusTypeDef MIPI_DSI_Start(DSI_HandleTypeDef *hdsi);", functionName: "MIPI_DSI_Start", body: ["  if (!hdsi) return HAL_ERROR;", "  return HAL_DSI_Start(hdsi);"], specs: [["INTERFACE", "MIPI-DSI Host"], ["MODE", "Video / Command"], ["LANES", "1 or 2, MCU dependent"]], tags: ["DSI HOST", "D-PHY", "LTDC"], guideUrl: "https://www.st.com/resource/en/application_note/an4860-introduction-to-dsi-host-on-stm32-mcus-and-mpus-stmicroelectronics.pdf" },
  { name: "FMC 8080/6800 LCD", slug: "fmc_8080_6800_lcd", category: "Display", summaryEn: "Write commands and pixels to a GRAM display through an FMC memory-mapped bus.", summaryKo: "FMC 메모리 매핑 버스로 GRAM 디스플레이에 명령과 픽셀을 씁니다.", applicationEn: "Intel 8080 and Motorola 6800 command displays with 8- or 16-bit buses.", applicationKo: "8·16비트 버스의 Intel 8080 및 Motorola 6800 명령형 디스플레이.", signature: "volatile uint16_t *command, volatile uint16_t *data, uint16_t command_value, uint16_t data_value", declaration: "HAL_StatusTypeDef FMC_LCD_Write(volatile uint16_t *command, volatile uint16_t *data, uint16_t command_value, uint16_t data_value);", functionName: "FMC_LCD_Write", body: ["  if (!command || !data) return HAL_ERROR;", "  *command = command_value;", "  *data = data_value;", "  return HAL_OK;"], specs: [["INTERFACE", "Intel 8080 / Motorola 6800"], ["BUS", "FMC 8 / 16-bit"], ["BUFFER", "Display GRAM"]], tags: ["FMC", "MEMORY MAPPED", "GRAM"], guideUrl: "https://support.touchgfx.com/docs/development/touchgfx-hal-development/scenarios/scenarios-fmc" },
  { name: "SPI LCD", slug: "spi_lcd", category: "Display", summaryEn: "Send command or pixel blocks to a controller-based display over SPI.", summaryKo: "SPI로 컨트롤러형 디스플레이에 명령 또는 픽셀 블록을 전송합니다.", applicationEn: "Small low-pin-count TFT, LCD and OLED modules with integrated GRAM.", applicationKo: "GRAM이 내장된 소형 저핀수 TFT, LCD 및 OLED 모듈.", signature: "SPI_HandleTypeDef *hspi, GPIO_TypeDef *dc_port, uint16_t dc_pin, const uint8_t *data, uint16_t size", declaration: "HAL_StatusTypeDef SPI_LCD_Send(SPI_HandleTypeDef *hspi, GPIO_TypeDef *dc_port, uint16_t dc_pin, const uint8_t *data, uint16_t size);", functionName: "SPI_LCD_Send", body: ["  if (!hspi || !dc_port || !data || size == 0U) return HAL_ERROR;", "  HAL_GPIO_WritePin(dc_port, dc_pin, GPIO_PIN_SET);", "  return HAL_SPI_Transmit(hspi, (uint8_t *)data, size, 100U);"], specs: [["INTERFACE", "SPI / MIPI-DBI Type C"], ["WIRES", "3 or 4 wire"], ["BUFFER", "Display GRAM"]], tags: ["SPI", "LOW PIN COUNT", "PARTIAL FRAMEBUFFER"], guideUrl: "https://support.touchgfx.com/docs/development/hardware-selection/hardware-components/hardware-selection-display" },
  { name: "Quad-SPI LCD", slug: "quad_spi_lcd", category: "Display", summaryEn: "Initialize a custom Quad-SPI display transport for higher serial bandwidth.", summaryKo: "더 높은 직렬 대역폭을 위한 사용자 정의 Quad-SPI 디스플레이 전송을 초기화합니다.", applicationEn: "Controller-based panels exposing QSPI-compatible command and pixel transfers.", applicationKo: "QSPI 호환 명령 및 픽셀 전송을 제공하는 컨트롤러형 패널.", signature: "QSPI_HandleTypeDef *hqspi", declaration: "HAL_StatusTypeDef QuadSPI_LCD_Start(QSPI_HandleTypeDef *hqspi);", functionName: "QuadSPI_LCD_Start", body: ["  if (!hqspi) return HAL_ERROR;", "  return HAL_QSPI_Init(hqspi);"], specs: [["INTERFACE", "Quad-SPI custom"], ["LANES", "1 / 4 data lines"], ["BUFFER", "Display GRAM"]], tags: ["QSPI", "CUSTOM DISPLAY", "HIGHER BANDWIDTH"], guideUrl: "https://support.touchgfx.com/docs/development/hardware-selection/hardware-components/hardware-selection-display" },
  { name: "eDP LCD Bridge", slug: "edp_lcd_bridge", category: "Display", summaryEn: "Program an external RGB or DSI-to-eDP bridge over its control bus.", summaryKo: "제어 버스로 외부 RGB 또는 DSI-eDP 브리지를 설정합니다.", applicationEn: "Laptop-style embedded DisplayPort panels connected through a board-specific bridge.", applicationKo: "보드 전용 브리지를 통해 연결하는 노트북형 Embedded DisplayPort 패널.", signature: "I2C_HandleTypeDef *hi2c, uint16_t address, uint8_t reg, uint8_t value", declaration: "HAL_StatusTypeDef EDP_BridgeWrite(I2C_HandleTypeDef *hi2c, uint16_t address, uint8_t reg, uint8_t value);", functionName: "EDP_BridgeWrite", body: ["  if (!hi2c) return HAL_ERROR;", "  return HAL_I2C_Mem_Write(hi2c, address, reg, I2C_MEMADD_SIZE_8BIT, &value, 1U, 100U);"], specs: [["PANEL", "Embedded DisplayPort"], ["SOURCE", "LTDC RGB or DSI"], ["BRIDGE", "External / I2C control"]], tags: ["EDP", "EXTERNAL BRIDGE", "BOARD SPECIFIC"], guideUrl: "https://support.touchgfx.com/docs/development/hardware-selection/hardware-components/hardware-selection-display" },
  { name: "HDMI Display Bridge", slug: "hdmi_display_bridge", category: "Display", summaryEn: "Configure an external RGB or DSI-to-HDMI transmitter over I²C.", summaryKo: "I²C로 외부 RGB 또는 DSI-HDMI 송신기를 설정합니다.", applicationEn: "Embedded HDMI monitors and service displays driven through a board-specific bridge.", applicationKo: "보드 전용 브리지를 통해 구동하는 임베디드 HDMI 모니터 및 서비스 화면.", signature: "I2C_HandleTypeDef *hi2c, uint16_t address, uint8_t reg, uint8_t value", declaration: "HAL_StatusTypeDef HDMI_BridgeWrite(I2C_HandleTypeDef *hi2c, uint16_t address, uint8_t reg, uint8_t value);", functionName: "HDMI_BridgeWrite", body: ["  if (!hi2c) return HAL_ERROR;", "  return HAL_I2C_Mem_Write(hi2c, address, reg, I2C_MEMADD_SIZE_8BIT, &value, 1U, 100U);"], specs: [["OUTPUT", "HDMI transmitter"], ["SOURCE", "LTDC RGB or DSI"], ["BRIDGE", "External / I2C control"]], tags: ["HDMI", "EXTERNAL BRIDGE", "BOARD SPECIFIC"], guideUrl: "https://www.st.com/content/st_com/en/ecosystems/stm32-graphic-user-interface/stm32-mcu.html" },
  { name: "I²C Display", slug: "i2c_display", category: "Display", summaryEn: "Send compact display-controller commands over an I²C control bus.", summaryKo: "I²C 제어 버스로 소형 디스플레이 컨트롤러 명령을 전송합니다.", applicationEn: "Small monochrome LCD and OLED status panels with integrated display RAM.", applicationKo: "디스플레이 RAM이 내장된 소형 단색 LCD 및 OLED 상태 패널.", signature: "I2C_HandleTypeDef *hi2c, uint16_t address, uint8_t command", declaration: "HAL_StatusTypeDef I2C_DisplayCommand(I2C_HandleTypeDef *hi2c, uint16_t address, uint8_t command);", functionName: "I2C_DisplayCommand", body: ["  if (!hi2c) return HAL_ERROR;", "  return HAL_I2C_Master_Transmit(hi2c, address, &command, 1U, 100U);"], specs: [["INTERFACE", "I2C control"], ["TARGET", "Small LCD / OLED"], ["BUFFER", "Display RAM"]], tags: ["I2C", "LOW POWER", "STATUS DISPLAY"], guideUrl: "https://support.touchgfx.com/docs/development/hardware-selection/hardware-components/hardware-selection-display" },
  { name: "GPIO", slug: "gpio", category: "System", summaryEn: "Read and write digital pins through STM32 HAL.", summaryKo: "STM32 HAL로 디지털 핀을 읽고 씁니다.", applicationEn: "LEDs, chip selects, enables and digital status signals.", applicationKo: "LED, 칩 선택, 활성화 및 디지털 상태 신호.", signature: "GPIO_TypeDef *port, uint16_t pin, GPIO_PinState state", declaration: "HAL_StatusTypeDef GPIO_Write(GPIO_TypeDef *port, uint16_t pin, GPIO_PinState state);", functionName: "GPIO_Write", body: ["  if (port == NULL) return HAL_ERROR;", "  HAL_GPIO_WritePin(port, pin, state);", "  return HAL_OK;"], guideUrl: wiki + "Getting_started_with_GPIO" },
  { name: "EXTI", slug: "exti", category: "System", summaryEn: "Enable and handle external GPIO interrupts.", summaryKo: "외부 GPIO 인터럽트를 활성화하고 처리합니다.", applicationEn: "Buttons, alarms, data-ready and wake-up events.", applicationKo: "버튼, 알람, 데이터 준비 및 웨이크업 이벤트.", signature: "IRQn_Type irq", declaration: "HAL_StatusTypeDef EXTI_Enable(IRQn_Type irq);", functionName: "EXTI_Enable", body: ["  HAL_NVIC_EnableIRQ(irq);", "  return HAL_OK;"], guideUrl: wiki + "Getting_started_with_EXTI" },
  { name: "PWR", slug: "pwr", category: "System", summaryEn: "Enter a low-power sleep state through HAL PWR.", summaryKo: "HAL PWR로 저전력 절전 상태에 진입합니다.", applicationEn: "Battery-powered idle and event-driven wake workflows.", applicationKo: "배터리 구동 유휴 및 이벤트 기반 웨이크업.", signature: "void", declaration: "HAL_StatusTypeDef PWR_EnterSleep(void);", functionName: "PWR_EnterSleep", body: ["  HAL_PWR_EnterSLEEPMode(PWR_MAINREGULATOR_ON, PWR_SLEEPENTRY_WFI);", "  return HAL_OK;"], guideUrl: wiki + "Getting_started_with_PWR" },
  { name: "DMA", slug: "dma", category: "System", summaryEn: "Move data between peripherals and memory with interrupts.", summaryKo: "인터럽트로 주변장치와 메모리 사이에서 데이터를 이동합니다.", applicationEn: "Low-CPU data acquisition and communication transfers.", applicationKo: "CPU 부하가 낮은 데이터 수집 및 통신 전송.", signature: "DMA_HandleTypeDef *hdma, uint32_t src, uint32_t dst, uint32_t size", declaration: "HAL_StatusTypeDef DMA_Start(DMA_HandleTypeDef *hdma, uint32_t src, uint32_t dst, uint32_t size);", functionName: "DMA_Start", body: ["  if (!hdma || size == 0U) return HAL_ERROR;", "  return HAL_DMA_Start_IT(hdma, src, dst, size);"], guideUrl: wiki + "Getting_started_with_DMA" },
  { name: "GPDMA", slug: "gpdma", category: "System", summaryEn: "Start a general-purpose DMA transfer.", summaryKo: "범용 DMA 전송을 시작합니다.", applicationEn: "Linked or high-throughput transfers on supported families.", applicationKo: "지원 제품군의 연결형 또는 고속 전송.", signature: "DMA_HandleTypeDef *hdma, uint32_t src, uint32_t dst, uint32_t size", declaration: "HAL_StatusTypeDef GPDMA_Start(DMA_HandleTypeDef *hdma, uint32_t src, uint32_t dst, uint32_t size);", functionName: "GPDMA_Start", body: ["  if (!hdma || size == 0U) return HAL_ERROR;", "  return HAL_DMA_Start_IT(hdma, src, dst, size);"], guideUrl: wiki + "Getting_started_with_GPDMA_and_Linked-list" },
  { name: "ADC", slug: "adc", category: "Analog", summaryEn: "Start an analog-to-digital conversion.", summaryKo: "아날로그-디지털 변환을 시작합니다.", applicationEn: "Sensor acquisition, voltage monitoring and analog measurements.", applicationKo: "센서 수집, 전압 감시 및 아날로그 측정.", signature: "ADC_HandleTypeDef *hadc", declaration: "HAL_StatusTypeDef ADC_Start(ADC_HandleTypeDef *hadc);", functionName: "ADC_Start", body: ["  if (!hadc) return HAL_ERROR;", "  return HAL_ADC_Start(hadc);"], guideUrl: wiki + "Getting_started_with_ADC" },
  { name: "DAC", slug: "dac", category: "Analog", summaryEn: "Start a DAC channel and set a 12-bit value.", summaryKo: "DAC 채널을 시작하고 12비트 값을 설정합니다.", applicationEn: "Analog setpoints, waveform generation and bias control.", applicationKo: "아날로그 설정값, 파형 생성 및 바이어스 제어.", signature: "DAC_HandleTypeDef *hdac, uint32_t channel, uint32_t value", declaration: "HAL_StatusTypeDef DAC_Write(DAC_HandleTypeDef *hdac, uint32_t channel, uint32_t value);", functionName: "DAC_Write", body: ["  if (!hdac || value > 4095U) return HAL_ERROR;", "  if (HAL_DAC_Start(hdac, channel) != HAL_OK) return HAL_ERROR;", "  return HAL_DAC_SetValue(hdac, channel, DAC_ALIGN_12B_R, value);"], guideUrl: wiki + "Getting_started_with_DAC" },
  { name: "TIM Base", slug: "tim_base", category: "Timing", summaryEn: "Start a base timer with period interrupts.", summaryKo: "주기 인터럽트가 있는 기본 타이머를 시작합니다.", applicationEn: "Periodic scheduling, timeouts and application ticks.", applicationKo: "주기 스케줄링, 타임아웃 및 애플리케이션 틱.", signature: "TIM_HandleTypeDef *htim", declaration: "HAL_StatusTypeDef TIM_BaseStart(TIM_HandleTypeDef *htim);", functionName: "TIM_BaseStart", body: ["  if (!htim) return HAL_ERROR;", "  return HAL_TIM_Base_Start_IT(htim);"], guideUrl: wiki + "Getting_started_with_TIM" },
  { name: "HRTIM", slug: "hrtim", category: "Timing", summaryEn: "Start high-resolution timer waveform counters.", summaryKo: "고해상도 타이머 파형 카운터를 시작합니다.", applicationEn: "Digital power, resonant converters and precision PWM.", applicationKo: "디지털 전원, 공진 컨버터 및 정밀 PWM.", signature: "HRTIM_HandleTypeDef *hhrtim, uint32_t timers", declaration: "HAL_StatusTypeDef HRTIM_Start(HRTIM_HandleTypeDef *hhrtim, uint32_t timers);", functionName: "HRTIM_Start", body: ["  if (!hhrtim) return HAL_ERROR;", "  return HAL_HRTIM_WaveformCounterStart(hhrtim, timers);"], guideUrl: wiki + "Getting_started_with_HRTIM" },
  { name: "I3C", slug: "i3c", category: "Communication", summaryEn: "Initialize an I3C controller on supported devices.", summaryKo: "지원 장치에서 I3C 컨트롤러를 초기화합니다.", applicationEn: "High-speed sensor buses with dynamic addressing.", applicationKo: "동적 주소 지정을 사용하는 고속 센서 버스.", signature: "I3C_HandleTypeDef *hi3c", declaration: "HAL_StatusTypeDef I3C_Start(I3C_HandleTypeDef *hi3c);", functionName: "I3C_Start", body: ["  if (!hi3c) return HAL_ERROR;", "  return HAL_I3C_Init(hi3c);"], guideUrl: wiki + "Getting_started_with_I3C" },
  { name: "CAN", slug: "can", category: "Communication", summaryEn: "Start a classic CAN controller.", summaryKo: "Classic CAN 컨트롤러를 시작합니다.", applicationEn: "Automotive, industrial and distributed control networks.", applicationKo: "자동차, 산업 및 분산 제어 네트워크.", signature: "CAN_HandleTypeDef *hcan", declaration: "HAL_StatusTypeDef CAN_Start(CAN_HandleTypeDef *hcan);", functionName: "CAN_Start", body: ["  if (!hcan) return HAL_ERROR;", "  return HAL_CAN_Start(hcan);"] },
  { name: "FDCAN", slug: "fdcan", category: "Communication", summaryEn: "Start a CAN FD controller.", summaryKo: "CAN FD 컨트롤러를 시작합니다.", applicationEn: "Higher-payload and higher-rate CAN networks.", applicationKo: "더 큰 페이로드와 높은 속도의 CAN 네트워크.", signature: "FDCAN_HandleTypeDef *hfdcan", declaration: "HAL_StatusTypeDef FDCAN_Start(FDCAN_HandleTypeDef *hfdcan);", functionName: "FDCAN_Start", body: ["  if (!hfdcan) return HAL_ERROR;", "  return HAL_FDCAN_Start(hfdcan);"] },
  { name: "RTC", slug: "rtc", category: "Timing", summaryEn: "Read coherent RTC time and date values.", summaryKo: "일관된 RTC 시간과 날짜 값을 읽습니다.", applicationEn: "Calendars, timestamps, alarms and low-power wakeups.", applicationKo: "달력, 타임스탬프, 알람 및 저전력 웨이크업.", signature: "RTC_HandleTypeDef *hrtc, RTC_TimeTypeDef *time, RTC_DateTypeDef *date", declaration: "HAL_StatusTypeDef RTC_Read(RTC_HandleTypeDef *hrtc, RTC_TimeTypeDef *time, RTC_DateTypeDef *date);", functionName: "RTC_Read", body: ["  if (!hrtc || !time || !date) return HAL_ERROR;", "  if (HAL_RTC_GetTime(hrtc, time, RTC_FORMAT_BIN) != HAL_OK) return HAL_ERROR;", "  return HAL_RTC_GetDate(hrtc, date, RTC_FORMAT_BIN);"], guideUrl: wiki + "Getting_started_with_RTC" },
  { name: "IWDG", slug: "iwdg", category: "System", summaryEn: "Refresh the independent watchdog.", summaryKo: "독립형 워치독을 갱신합니다.", applicationEn: "Recovery from firmware stalls and safety supervision.", applicationKo: "펌웨어 정지 복구 및 안전 감시.", signature: "IWDG_HandleTypeDef *hiwdg", declaration: "HAL_StatusTypeDef IWDG_Service(IWDG_HandleTypeDef *hiwdg);", functionName: "IWDG_Service", body: ["  if (!hiwdg) return HAL_ERROR;", "  return HAL_IWDG_Refresh(hiwdg);"], guideUrl: wiki + "Getting_started_with_WDG" },
  { name: "USB Device", slug: "usb_device", category: "Connectivity", summaryEn: "Initialize the generated USB device stack.", summaryKo: "생성된 USB Device 스택을 초기화합니다.", applicationEn: "CDC, HID, MSC and other USB device classes.", applicationKo: "CDC, HID, MSC 및 기타 USB Device 클래스.", signature: "void", declaration: "HAL_StatusTypeDef USB_DeviceStart(void);", functionName: "USB_DeviceStart", body: ["  MX_USB_DEVICE_Init();", "  return HAL_OK;"], guideUrl: wiki + "Introduction_to_USB_with_STM32" },
  { name: "USB Host", slug: "usb_host", category: "Connectivity", summaryEn: "Initialize the generated USB host stack.", summaryKo: "생성된 USB Host 스택을 초기화합니다.", applicationEn: "Host-mode storage, HID and supported USB classes.", applicationKo: "Host 모드 저장장치, HID 및 지원 USB 클래스.", signature: "void", declaration: "HAL_StatusTypeDef USB_HostStart(void);", functionName: "USB_HostStart", body: ["  MX_USB_HOST_Init();", "  return HAL_OK;"], guideUrl: wiki + "Introduction_to_USB_with_STM32" },
  { name: "SDMMC", slug: "sdmmc", category: "Storage", summaryEn: "Read SD card blocks through HAL SD.", summaryKo: "HAL SD로 SD 카드 블록을 읽습니다.", applicationEn: "Data logging, removable media and FAT file systems.", applicationKo: "데이터 로깅, 이동식 미디어 및 FAT 파일 시스템.", signature: "SD_HandleTypeDef *hsd, uint8_t *data, uint32_t block, uint32_t count", declaration: "HAL_StatusTypeDef SDMMC_Read(SD_HandleTypeDef *hsd, uint8_t *data, uint32_t block, uint32_t count);", functionName: "SDMMC_Read", body: ["  if (!hsd || !data || count == 0U) return HAL_ERROR;", "  return HAL_SD_ReadBlocks(hsd, data, block, count, 1000U);"] },
  { name: "QSPI", slug: "qspi", category: "Storage", summaryEn: "Initialize a Quad-SPI external memory interface.", summaryKo: "Quad-SPI 외부 메모리 인터페이스를 초기화합니다.", applicationEn: "External NOR flash, assets and execute-in-place designs.", applicationKo: "외부 NOR 플래시, 자산 및 XIP 설계.", signature: "QSPI_HandleTypeDef *hqspi", declaration: "HAL_StatusTypeDef QSPI_Start(QSPI_HandleTypeDef *hqspi);", functionName: "QSPI_Start", body: ["  if (!hqspi) return HAL_ERROR;", "  return HAL_QSPI_Init(hqspi);"] },
  { name: "OSPI", slug: "ospi", category: "Storage", summaryEn: "Initialize an Octo-SPI external memory interface.", summaryKo: "Octo-SPI 외부 메모리 인터페이스를 초기화합니다.", applicationEn: "High-throughput external flash and PSRAM.", applicationKo: "고속 외부 플래시 및 PSRAM.", signature: "OSPI_HandleTypeDef *hospi", declaration: "HAL_StatusTypeDef OSPI_Start(OSPI_HandleTypeDef *hospi);", functionName: "OSPI_Start", body: ["  if (!hospi) return HAL_ERROR;", "  return HAL_OSPI_Init(hospi);"] },
  { name: "SAI", slug: "sai", category: "Audio", summaryEn: "Transmit audio samples through SAI.", summaryKo: "SAI를 통해 오디오 샘플을 전송합니다.", applicationEn: "Codecs, multichannel digital audio and microphones.", applicationKo: "코덱, 다채널 디지털 오디오 및 마이크.", signature: "SAI_HandleTypeDef *hsai, uint8_t *data, uint16_t size", declaration: "HAL_StatusTypeDef SAI_Send(SAI_HandleTypeDef *hsai, uint8_t *data, uint16_t size);", functionName: "SAI_Send", body: ["  if (!hsai || !data || size == 0U) return HAL_ERROR;", "  return HAL_SAI_Transmit(hsai, data, size, 100U);"] },
  { name: "I2S", slug: "i2s", category: "Audio", summaryEn: "Transmit digital audio samples through I2S.", summaryKo: "I2S를 통해 디지털 오디오 샘플을 전송합니다.", applicationEn: "Stereo codecs, DACs and audio stream links.", applicationKo: "스테레오 코덱, DAC 및 오디오 스트림 링크.", signature: "I2S_HandleTypeDef *hi2s, uint16_t *data, uint16_t size", declaration: "HAL_StatusTypeDef I2S_Send(I2S_HandleTypeDef *hi2s, uint16_t *data, uint16_t size);", functionName: "I2S_Send", body: ["  if (!hi2s || !data || size == 0U) return HAL_ERROR;", "  return HAL_I2S_Transmit(hi2s, data, size, 100U);"] },
  { name: "DCMI", slug: "dcmi", category: "Imaging", summaryEn: "Start camera capture into a DMA framebuffer.", summaryKo: "DMA 프레임버퍼로 카메라 캡처를 시작합니다.", applicationEn: "Parallel camera sensors and image acquisition.", applicationKo: "병렬 카메라 센서 및 이미지 수집.", signature: "DCMI_HandleTypeDef *hdcmi, uint32_t address, uint32_t words", declaration: "HAL_StatusTypeDef DCMI_Capture(DCMI_HandleTypeDef *hdcmi, uint32_t address, uint32_t words);", functionName: "DCMI_Capture", body: ["  if (!hdcmi || address == 0U || words == 0U) return HAL_ERROR;", "  return HAL_DCMI_Start_DMA(hdcmi, DCMI_MODE_CONTINUOUS, address, words);"] },
  { name: "RNG", slug: "rng", category: "Security", summaryEn: "Generate a hardware random word.", summaryKo: "하드웨어 난수 값을 생성합니다.", applicationEn: "Nonces, keys, randomized protocols and security services.", applicationKo: "Nonce, 키, 무작위 프로토콜 및 보안 서비스.", signature: "RNG_HandleTypeDef *hrng, uint32_t *value", declaration: "HAL_StatusTypeDef RNG_Read(RNG_HandleTypeDef *hrng, uint32_t *value);", functionName: "RNG_Read", body: ["  if (!hrng || !value) return HAL_ERROR;", "  return HAL_RNG_GenerateRandomNumber(hrng, value);"] },
  { name: "CRC", slug: "crc", category: "Security", summaryEn: "Calculate a hardware-assisted CRC over words.", summaryKo: "워드 데이터의 하드웨어 가속 CRC를 계산합니다.", applicationEn: "Packet validation, storage integrity and firmware checks.", applicationKo: "패킷 검증, 저장 무결성 및 펌웨어 검사.", signature: "CRC_HandleTypeDef *hcrc, uint32_t *data, uint32_t words, uint32_t *result", declaration: "HAL_StatusTypeDef CRC_Calculate(CRC_HandleTypeDef *hcrc, uint32_t *data, uint32_t words, uint32_t *result);", functionName: "CRC_Calculate", body: ["  if (!hcrc || !data || !result || words == 0U) return HAL_ERROR;", "  *result = HAL_CRC_Calculate(hcrc, data, words);", "  return HAL_OK;"] },
];
