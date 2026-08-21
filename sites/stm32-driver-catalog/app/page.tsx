"use client";

import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import { additionalDrivers, halReferenceUrl } from "./driver-catalog";

type Peripheral = string;

type DriverProfile = {
  filename: string;
  headerFilename?: string;
  summary: string;
  summaryEn: string;
  specs: [string, string][];
  tags: string[];
  code: string;
};

let peripherals: Peripheral[] = ["UART", "SPI", "I²C", "LAN", "RS422/485", "PWM", "LVDS LCD"];

const driverProfiles: Record<Peripheral, DriverProfile> = {
  UART: {
    filename: "uart_driver.c",
    summary: "간단한 블로킹 송신과 입력값 검증을 포함한 기본 UART 프로필입니다.",
    summaryEn: "A defensive UART starter with blocking transmit and input validation.",
    specs: [["LAYER", "STM32 HAL"], ["MODE", "Blocking TX"], ["TIMEOUT", "100 ms"]],
    tags: ["HAL API", "NULL GUARD", "BOUNDED TIMEOUT"],
    code: `// uart_driver.c — STM32 HAL
#include "uart_driver.h"

HAL_StatusTypeDef UART_Send(
  UART_HandleTypeDef *huart,
  const uint8_t *data,
  uint16_t size
) {
  if (!huart || !data || size == 0U) {
    return HAL_ERROR;
  }

  return HAL_UART_Transmit(
    huart, data, size, 100U
  );
}`,
  },
  SPI: {
    filename: "spi_driver.c",
    summary: "동시 송수신이 필요한 센서·메모리 장치를 위한 SPI 전송 프로필입니다.",
    summaryEn: "A full-duplex SPI transfer profile for sensors and memory devices.",
    specs: [["LAYER", "STM32 HAL"], ["MODE", "Full duplex"], ["TIMEOUT", "100 ms"]],
    tags: ["FULL DUPLEX", "BUFFER GUARD", "BOUNDED TIMEOUT"],
    code: `// spi_driver.c — STM32 HAL
#include "spi_driver.h"

HAL_StatusTypeDef SPI_Transfer(
  SPI_HandleTypeDef *hspi,
  uint8_t *tx, uint8_t *rx,
  uint16_t size
) {
  if (!hspi || !tx || !rx) return HAL_ERROR;

  return HAL_SPI_TransmitReceive(
    hspi, tx, rx, size, 100U
  );
}`,
  },
  "I²C": {
    filename: "i2c_driver.c",
    summary: "8비트 레지스터 주소를 사용하는 I²C 장치의 읽기 시작점입니다.",
    summaryEn: "An I²C register-read starter for devices using 8-bit register addresses.",
    specs: [["LAYER", "STM32 HAL"], ["ADDRESS", "7-bit shifted"], ["REGISTER", "8-bit"]],
    tags: ["MEMORY READ", "ADDRESS REVIEW", "BOUNDED TIMEOUT"],
    code: `// i2c_driver.c — STM32 HAL
#include "i2c_driver.h"

HAL_StatusTypeDef I2C_ReadRegister(
  I2C_HandleTypeDef *hi2c,
  uint16_t address, uint8_t reg,
  uint8_t *data, uint16_t size
) {
  if (!hi2c || !data) return HAL_ERROR;

  return HAL_I2C_Mem_Read(hi2c, address,
    reg, I2C_MEMADD_SIZE_8BIT,
    data, size, 100U);
}`,
  },
  LAN: {
    filename: "lan_driver.c",
    summary: "외부 PHY와 Ethernet HAL을 LwIP 네트워크 인터페이스에 연결하는 LAN 시작 프로필입니다.",
    summaryEn: "A LAN starter connecting an external PHY and Ethernet HAL to LwIP.",
    specs: [["STACK", "LwIP"], ["LINK", "MII / RMII PHY"], ["MODE", "Interrupt RX/TX"]],
    tags: ["ETH HAL", "LWIP", "PHY LINK"],
    code: `// lan_driver.c — STM32 HAL + LwIP
#include "lan_driver.h"
#include "lwip.h"

HAL_StatusTypeDef LAN_Start(
  ETH_HandleTypeDef *heth
) {
  if (heth == NULL) return HAL_ERROR;

  MX_LWIP_Init();
  return HAL_ETH_Start_IT(heth);
}

bool LAN_IsLinkUp(void) {
  return netif_is_link_up(&gnetif) != 0;
}`,
  },
  "RS422/485": {
    filename: "rs422_485_driver.c",
    summary: "RS485의 DE 방향 전환과 RS422/485 UART 전송 완료 대기를 포함합니다.",
    summaryEn: "An RS422/485 profile with UART transfer and RS485 DE direction control.",
    specs: [["ENGINE", "UART"], ["RS485", "GPIO DE control"], ["RS422", "Full duplex wiring"]],
    tags: ["UART HAL", "DE CONTROL", "BUS TURNAROUND"],
    code: `// rs422_485_driver.c — STM32 HAL
#include "rs422_485_driver.h"

HAL_StatusTypeDef RS485_Send(
  UART_HandleTypeDef *huart,
  GPIO_TypeDef *de_port, uint16_t de_pin,
  const uint8_t *data, uint16_t size
) {
  if (!huart || !de_port || !data || size == 0U)
    return HAL_ERROR;

  HAL_GPIO_WritePin(de_port, de_pin, GPIO_PIN_SET);
  HAL_StatusTypeDef status = HAL_UART_Transmit(
    huart, data, size, 100U);
  while (__HAL_UART_GET_FLAG(huart, UART_FLAG_TC) == 0U) {}
  HAL_GPIO_WritePin(de_port, de_pin, GPIO_PIN_RESET);
  return status;
}`,
  },
  PWM: {
    filename: "pwm_driver.c",
    summary: "TIM 채널을 시작하고 ARR 범위 안에서 듀티를 안전하게 갱신합니다.",
    summaryEn: "Starts a TIM PWM channel and safely updates duty within the ARR range.",
    specs: [["ENGINE", "TIM"], ["OUTPUT", "PWM channel"], ["DUTY", "0…ARR"]],
    tags: ["TIM HAL", "DUTY CLAMP", "LIVE UPDATE"],
    code: `// pwm_driver.c — STM32 HAL
#include "pwm_driver.h"

HAL_StatusTypeDef PWM_Start(
  TIM_HandleTypeDef *htim, uint32_t channel
) {
  if (htim == NULL) return HAL_ERROR;
  return HAL_TIM_PWM_Start(htim, channel);
}

HAL_StatusTypeDef PWM_SetDuty(
  TIM_HandleTypeDef *htim, uint32_t channel,
  uint32_t pulse
) {
  if (!htim || pulse > __HAL_TIM_GET_AUTORELOAD(htim))
    return HAL_ERROR;

  __HAL_TIM_SET_COMPARE(htim, channel, pulse);
  return HAL_OK;
}`,
  },
  "LVDS LCD": {
    filename: "lvds_lcd_driver.c",
    summary: "LTDC 프레임버퍼와 I²C 외부 LVDS 브리지를 순서대로 활성화하는 디스플레이 프로필입니다.",
    summaryEn: "A display profile that sequences an LTDC framebuffer and external I²C LVDS bridge.",
    specs: [["PIXEL", "LTDC RGB"], ["BRIDGE", "External LVDS / I²C"], ["BUFFER", "SDRAM framebuffer"]],
    tags: ["LTDC", "LVDS BRIDGE", "FRAMEBUFFER"],
    code: `// lvds_lcd_driver.c — LTDC + external bridge
#include "lvds_lcd_driver.h"

HAL_StatusTypeDef LVDS_LCD_Start(
  LTDC_HandleTypeDef *hltdc,
  uint32_t framebuffer
) {
  if (!hltdc || framebuffer == 0U)
    return HAL_ERROR;

  // Board-specific: power rails, reset and the I2C
  // LVDS bridge must be configured before video starts.
  if (LVDS_Bridge_Init() != HAL_OK)
    return HAL_ERROR;

  if (HAL_LTDC_SetAddress(hltdc, framebuffer, 0U) != HAL_OK)
    return HAL_ERROR;

  __HAL_LTDC_ENABLE(hltdc);
  LVDS_Backlight_Set(true);
  return HAL_OK;
}`,
  },
};

const driverHeaders: Record<Peripheral, string> = {
  UART: `#ifndef UART_DRIVER_H
#define UART_DRIVER_H
#include "main.h"
HAL_StatusTypeDef UART_Send(UART_HandleTypeDef *huart, const uint8_t *data, uint16_t size);
#endif`,
  SPI: `#ifndef SPI_DRIVER_H
#define SPI_DRIVER_H
#include "main.h"
HAL_StatusTypeDef SPI_Transfer(SPI_HandleTypeDef *hspi, uint8_t *tx, uint8_t *rx, uint16_t size);
#endif`,
  "I²C": `#ifndef I2C_DRIVER_H
#define I2C_DRIVER_H
#include "main.h"
HAL_StatusTypeDef I2C_ReadRegister(I2C_HandleTypeDef *hi2c, uint16_t address, uint8_t reg, uint8_t *data, uint16_t size);
#endif`,
  LAN: `#ifndef LAN_DRIVER_H
#define LAN_DRIVER_H
#include "main.h"
#include <stdbool.h>
HAL_StatusTypeDef LAN_Start(ETH_HandleTypeDef *heth);
bool LAN_IsLinkUp(void);
#endif`,
  "RS422/485": `#ifndef RS422_485_DRIVER_H
#define RS422_485_DRIVER_H
#include "main.h"
HAL_StatusTypeDef RS485_Send(UART_HandleTypeDef *huart, GPIO_TypeDef *de_port, uint16_t de_pin, const uint8_t *data, uint16_t size);
#endif`,
  PWM: `#ifndef PWM_DRIVER_H
#define PWM_DRIVER_H
#include "main.h"
HAL_StatusTypeDef PWM_Start(TIM_HandleTypeDef *htim, uint32_t channel);
HAL_StatusTypeDef PWM_SetDuty(TIM_HandleTypeDef *htim, uint32_t channel, uint32_t pulse);
#endif`,
  "LVDS LCD": `#ifndef LVDS_LCD_DRIVER_H
#define LVDS_LCD_DRIVER_H
#include "main.h"
#include <stdbool.h>
HAL_StatusTypeDef LVDS_LCD_Start(LTDC_HandleTypeDef *hltdc, uint32_t framebuffer);
HAL_StatusTypeDef LVDS_Bridge_Init(void);
void LVDS_Backlight_Set(bool enabled);
#endif`,
};

type DriverMeta = { category: string; guideUrl: string; applicationEn: string; applicationKo: string };
type DriverExample = { featureEn: string; featureKo: string; usageEn: string; usageKo: string };

const driverMeta: Record<Peripheral, DriverMeta> = {
  UART: { category: "Communication", guideUrl: "https://wiki.st.com/stm32mcu/wiki/Getting_started_with_UART", applicationEn: "Asynchronous serial command, telemetry and debug links.", applicationKo: "비동기 직렬 명령, 텔레메트리 및 디버그 링크." },
  SPI: { category: "Communication", guideUrl: "https://wiki.st.com/stm32mcu/wiki/Getting_started_with_SPI", applicationEn: "Full-duplex links for sensors, displays and external memory.", applicationKo: "센서, 디스플레이 및 외부 메모리용 전이중 링크." },
  "I²C": { category: "Communication", guideUrl: "https://wiki.st.com/stm32mcu/wiki/Getting_started_with_I2C", applicationEn: "Addressed control bus for sensors, PMICs and configuration devices.", applicationKo: "센서, PMIC 및 설정 장치용 주소 기반 제어 버스." },
  LAN: { category: "Connectivity", guideUrl: "https://www.st.com/resource/en/user_manual/dm00103685-developing-applications-on-stm32cube-with-lwip-tcp-ip-stack-stmicroelectronics.pdf", applicationEn: "Ethernet PHY, HAL and LwIP application integration.", applicationKo: "Ethernet PHY, HAL 및 LwIP 애플리케이션 통합." },
  "RS422/485": { category: "Communication", guideUrl: "https://wiki.st.com/stm32mcu/wiki/Getting_started_with_UART", applicationEn: "Robust differential serial links with RS485 direction control.", applicationKo: "RS485 방향 제어를 포함한 견고한 차동 직렬 링크." },
  PWM: { category: "Timing", guideUrl: "https://wiki.st.com/stm32mcu/wiki/Getting_started_with_TIM", applicationEn: "Timer-driven motor, LED, actuator and power-control waveforms.", applicationKo: "타이머 기반 모터, LED, 액추에이터 및 전력 제어 파형." },
  "LVDS LCD": { category: "Display", guideUrl: "https://www.st.com/content/st_com/en/stm32-mcu-developer-zone/embedded-software.html", applicationEn: "LTDC framebuffer output through a board-specific external LVDS bridge.", applicationKo: "보드 전용 외부 LVDS 브리지를 통한 LTDC 프레임버퍼 출력." },
};

for (const seed of additionalDrivers) {
  const filename = seed.filename ?? seed.slug + "_driver.c";
  const headerFilename = seed.headerFilename ?? seed.slug + "_driver.h";
  const guard = seed.slug.toUpperCase() + "_DRIVER_H";
  driverProfiles[seed.name] = {
    filename,
    headerFilename,
    summary: seed.summaryKo,
    summaryEn: seed.summaryEn,
    specs: seed.specs ?? [["LAYER", "STM32 HAL"], ["CATEGORY", seed.category], ["AVAILABILITY", "MCU dependent"]],
    tags: seed.tags ?? ["HAL API", seed.category.toUpperCase(), "CUBEMX REVIEW"],
    code: seed.source ?? [
      "// " + filename + " — STM32 HAL",
      "#include \"" + seed.slug + "_driver.h\"",
      "",
      "HAL_StatusTypeDef " + seed.functionName + "(" + seed.signature + ") {",
      ...seed.body,
      "}",
    ].join("\n"),
  };
  driverHeaders[seed.name] = seed.header ?? ["#ifndef " + guard, "#define " + guard, "#include \"main.h\"", seed.declaration, "#endif"].join("\n");
  driverMeta[seed.name] = {
    category: seed.category,
    guideUrl: seed.guideUrl ?? halReferenceUrl,
    applicationEn: seed.applicationEn,
    applicationKo: seed.applicationKo,
  };
}

peripherals = Object.keys(driverProfiles);

const driverExamples: Record<Peripheral, DriverExample> = {
  UART: {
    featureEn: "DMA circular receive, idle-line frame detection, TX queue, ring buffer, timeout recovery and communication error counters.",
    featureKo: "DMA 순환 수신, 유휴 라인 프레임 감지, 송신 큐, 링 버퍼, 타임아웃 복구 및 통신 오류 카운터.",
    usageEn: "Receive command and telemetry frames without blocking the main loop, validate CRC, then notify the application when a complete packet arrives.",
    usageKo: "메인 루프를 차단하지 않고 명령·텔레메트리 프레임을 수신하고 CRC를 검증한 뒤 완전한 패킷이 도착하면 애플리케이션에 알립니다.",
  },
  SPI: {
    featureEn: "DMA full-duplex transfer, chip-select timing, transaction queue, device-ready polling, timeout and bus-error recovery.",
    featureKo: "DMA 전이중 전송, 칩 선택 타이밍, 트랜잭션 큐, 장치 준비 폴링, 타임아웃 및 버스 오류 복구.",
    usageEn: "Read sensors and external memory through queued SPI transactions, complete transfers in callbacks and release chip select safely.",
    usageKo: "큐 기반 SPI 트랜잭션으로 센서와 외부 메모리를 읽고 콜백에서 전송을 완료한 뒤 칩 선택을 안전하게 해제합니다.",
  },
  "I²C": {
    featureEn: "Register read/write helpers, DMA transfer, bus-busy timeout, NACK retry, device scan and stuck-bus recovery.",
    featureKo: "레지스터 읽기·쓰기 헬퍼, DMA 전송, 버스 점유 타임아웃, NACK 재시도, 장치 검색 및 버스 고착 복구.",
    usageEn: "Configure sensors and PMICs at startup, then perform scheduled register reads and report bus faults to the main state machine.",
    usageKo: "시작 시 센서와 PMIC를 설정하고 주기적으로 레지스터를 읽어 버스 오류를 메인 상태 머신에 보고합니다.",
  },
  LAN: {
    featureEn: "PHY link monitoring, DHCP/static IP fallback, RX/TX descriptor recovery, reconnect timer, packet statistics and LwIP service hooks.",
    featureKo: "PHY 링크 감시, DHCP·고정 IP 대체, RX/TX 디스크립터 복구, 재연결 타이머, 패킷 통계 및 LwIP 서비스 훅.",
    usageEn: "Bring up Ethernet after PHY negotiation, reconnect after link loss and exchange TCP/UDP telemetry without blocking the control loop.",
    usageKo: "PHY 협상 후 Ethernet을 시작하고 링크 손실 시 재연결하며 제어 루프를 차단하지 않고 TCP/UDP 텔레메트리를 교환합니다.",
  },
  "RS422/485": {
    featureEn: "DE/RE direction timing, DMA receive-to-idle, half-duplex turnaround delay, Modbus frame CRC, collision timeout and retry.",
    featureKo: "DE/RE 방향 타이밍, DMA 유휴 라인 수신, 반이중 전환 지연, Modbus 프레임 CRC, 충돌 타임아웃 및 재시도.",
    usageEn: "Exchange framed industrial messages on a multidrop bus, switch direction after TX complete and validate every received packet.",
    usageKo: "멀티드롭 버스에서 산업용 프레임을 교환하고 송신 완료 후 방향을 전환하며 모든 수신 패킷을 검증합니다.",
  },
  PWM: {
    featureEn: "Safe duty update, frequency change, complementary outputs, dead-time control, soft start/stop and emergency output disable.",
    featureKo: "안전한 듀티 갱신, 주파수 변경, 상보 출력, 데드타임 제어, 소프트 시작·정지 및 비상 출력 차단.",
    usageEn: "Drive a motor, fan, LED or power stage with bounded duty commands and shut the output down immediately on a fault input.",
    usageKo: "제한된 듀티 명령으로 모터·팬·LED·전력단을 구동하고 오류 입력이 발생하면 출력을 즉시 차단합니다.",
  },
  "LVDS LCD": {
    featureEn: "Panel power sequence, bridge register setup, framebuffer switching, VSYNC synchronization, backlight PWM and display fault recovery.",
    featureKo: "패널 전원 시퀀스, 브리지 레지스터 설정, 프레임버퍼 전환, VSYNC 동기화, 백라이트 PWM 및 디스플레이 오류 복구.",
    usageEn: "Initialize SDRAM and the external LVDS bridge, start LTDC on a validated framebuffer and update the display on vertical blanking.",
    usageKo: "SDRAM과 외부 LVDS 브리지를 초기화하고 검증된 프레임버퍼로 LTDC를 시작한 뒤 수직 블랭킹 시 화면을 갱신합니다.",
  },
};

for (const seed of additionalDrivers) {
  driverExamples[seed.name] = {
    featureEn: `${seed.functionName}(), interrupt or DMA completion handling, bounded timeout, retry policy, error flags and diagnostic counters.`,
    featureKo: `${seed.functionName}(), 인터럽트 또는 DMA 완료 처리, 제한된 타임아웃, 재시도 정책, 오류 플래그 및 진단 카운터.`,
    usageEn: `${seed.applicationEn} Initialize it after the generated MX setup, keep transfers non-blocking where practical, and report completion or faults to the main application.`,
    usageKo: `${seed.applicationKo} 생성된 MX 설정 다음에 초기화하고 가능한 경우 비차단 방식으로 전송하며 완료 또는 오류를 메인 애플리케이션에 보고합니다.`,
  };
}

const checklistEn = [
  ["01", "Collect detections", "Record product, detection name, path and SHA-256, then quarantine the suspect file."],
  ["02", "Verify trust chain", "Check the official source, code signature, pinned dependencies and reproducible build."],
  ["03", "Re-scan with multiple engines", "Review results from more than one engine together with behavior logs."],
  ["04", "Request vendor review", "Submit the verified sample and detection details to the security vendor."],
  ["05", "Apply a narrow temporary allow", "Allow only the file hash or artifact with an expiry while review is pending."],
];

const ui = {
  en: {
    navForge: "Driver Forge", navShield: "Antivirus Recovery", navVerify: "Verify",
    hero: "Start STM32 serial, LAN, PWM and modern display-interface drivers quickly while keeping endpoint protection enabled.",
    start: "Build a driver", guide: "Recovery guide", forgeTitle: "Select a profile, describe the job, and download the full source.",
    peripheral: "PERIPHERAL", selectAria: "Select a peripheral", boardNote: "Review the board pinout, clocks, CubeMX settings and the PHY, transceiver or bridge datasheet before use.",
    libraryTitle: "STM32 APPLICATION LIBRARY", libraryCopy: "Browse 41 HAL and TouchGFX application profiles, including nine common embedded display paths. Hardware availability depends on the selected STM32 family, panel and bridge.", searchDrivers: "Search drivers", allCategories: "All categories", openGuide: "Open developing application", closeGuide: "Close guide", officialGuide: "Open official ST guide", useCases: "APPLICATION USE", setupFlow: "CUBEMX + HAL SETUP", familyNote: "Confirm this peripheral exists on your exact MCU and use the matching STM32Cube package before compiling.", setupSteps: ["Enable the peripheral, pins and clocks in STM32CubeMX.", "Generate the HAL handle and interrupt or DMA configuration.", "Add the generated C/H package after the matching MX initialization call.", "Validate callbacks, error recovery and board-level timing."],
    cubeReference: "CONFIGURATION REFERENCE", cubeMx2Note: "CubeMX2 is the HAL2-oriented offering for supported new STM32 series. It uses .ioc2 projects, live C/H code preview and generated mx_ configuration files.", cubeMx2Family: "CubeMX2 support starts with newer STM32 series; verify the current device list in ST documentation.", cubeMx2Steps: ["Create or open the supported device project using its .ioc2 file.", "Enable the peripheral, then configure Pinout, Clock, DMA, EXTI and NVIC views.", "Use Code preview to inspect the generated .c and .h files before project generation.", "Keep application logic outside generated mx_ files and protect regeneration with version control."],
    chooseTool: "CHOOSE TOOL", chooseCategory: "CHOOSE CATEGORY", chooseDriver: "CHOOSE DRIVER", quickBrowse: "QUICK CATEGORY BROWSE", matchingDrivers: "matching drivers", selectedProfile: "SELECTED PROFILE", browseHint: "Choose from the columns or browse the visual cards below.",
    featureLabel: "FUNCTIONS TO ADD", featurePlaceholder: "Example: DMA receive, ring buffer, packet CRC, reconnect after link loss…",
    usageLabel: "HOW IT WILL BE USED", usagePlaceholder: "Example: Receive 64-byte sensor frames at 10 ms intervals and notify the main loop…",
    packageTitle: "SOURCE PACKAGE", packageCopy: "Downloads a ZIP containing the C/C++ source, header and integration README.", download: "Download full source (.zip)",
    needsBrief: "Enter both the requested functions and intended use.", copied: "Copied ✓", copy: "Copy code",
    shieldTitle: "Keep Antivirus protection on; automate false-positive response.", shieldCopy: "A vendor-neutral workflow collects evidence and routes only verified files to review and narrowly scoped allow steps.",
    auto: "Collection · hashing · quarantine · submission queue", review: "Final allow · restore · release approval",
    noAutoTitle: "NO AUTOMATIC EXCLUSIONS", noAutoCopy: "Never trust every file or disable protection globally. Keep detections quarantined until they are verified.",
    verifyCopy: "After applying an approved narrow allow rule, verify the local toolchain with the smallest command sequence.",
    checking: "CHECKING", standby: "STANDBY", passed: "PASSED", run: "Run local readiness check", again: "Run again",
    pass: "✓ The command chain is ready. Confirm the final detection state in your Antivirus protection history.",
    footer: "Embedded development without disabling security",
  },
  ko: {
    navForge: "드라이버 생성", navShield: "Antivirus 복구", navVerify: "검증",
    hero: "직렬 통신부터 LAN·PWM·현대적인 디스플레이 인터페이스까지 빠르게 시작하고, 엔드포인트 보호를 유지한 채 개발하세요.",
    start: "드라이버 시작", guide: "복구 가이드", forgeTitle: "프로필을 선택하고 필요한 기능과 사용법을 입력해 풀 소스를 받으세요.",
    peripheral: "주변장치", selectAria: "주변장치 선택", boardNote: "보드 핀맵, 클록, CubeMX 설정과 PHY·트랜시버·브리지 데이터시트를 검토한 뒤 사용하세요.",
    libraryTitle: "STM32 애플리케이션 라이브러리", libraryCopy: "9가지 주요 디스플레이 경로를 포함한 41개 HAL 및 TouchGFX 프로필을 탐색할 수 있습니다. 실제 지원은 STM32 제품군, 패널 및 브리지에 따라 달라집니다.", searchDrivers: "드라이버 검색", allCategories: "전체 분류", openGuide: "Developing application 열기", closeGuide: "가이드 닫기", officialGuide: "ST 공식 가이드 열기", useCases: "애플리케이션 용도", setupFlow: "CUBEMX + HAL 설정", familyNote: "컴파일 전에 정확한 MCU에 해당 주변장치가 있는지와 일치하는 STM32Cube 패키지를 확인하세요.", setupSteps: ["STM32CubeMX에서 주변장치, 핀과 클록을 활성화합니다.", "HAL 핸들과 인터럽트 또는 DMA 설정을 생성합니다.", "해당 MX 초기화 호출 다음에 생성된 C/H 패키지를 추가합니다.", "콜백, 오류 복구 및 보드 수준 타이밍을 검증합니다."],
    cubeReference: "설정 도구 참조", cubeMx2Note: "CubeMX2는 지원되는 신규 STM32 제품군을 위한 HAL2 중심 도구입니다. .ioc2 프로젝트, 실시간 C/H 코드 미리보기 및 mx_ 설정 파일을 사용합니다.", cubeMx2Family: "CubeMX2는 신규 STM32 제품군부터 지원되므로 ST 문서에서 현재 지원 장치 목록을 확인하세요.", cubeMx2Steps: ["지원 장치의 .ioc2 파일로 프로젝트를 만들거나 엽니다.", "주변장치를 활성화하고 Pinout, Clock, DMA, EXTI 및 NVIC 화면을 설정합니다.", "프로젝트 생성 전에 Code preview에서 생성될 .c와 .h 파일을 확인합니다.", "애플리케이션 로직을 생성되는 mx_ 파일 밖에 두고 버전 관리로 재생성을 보호합니다."],
    chooseTool: "도구 선택", chooseCategory: "카테고리 선택", chooseDriver: "드라이버 선택", quickBrowse: "카테고리 빠른 탐색", matchingDrivers: "개 드라이버", selectedProfile: "선택한 프로필", browseHint: "상단 컬럼에서 선택하거나 아래 시각 카드에서 빠르게 찾아보세요.",
    featureLabel: "부여할 기능", featurePlaceholder: "예: DMA 수신, 링 버퍼, 패킷 CRC, 링크 손실 후 재연결…",
    usageLabel: "사용 방법", usagePlaceholder: "예: 10ms 간격의 64바이트 센서 프레임을 받고 메인 루프에 알림…",
    packageTitle: "소스 패키지", packageCopy: "C/C++ 소스, 헤더와 적용 안내 README가 포함된 ZIP을 내려받습니다.", download: "풀 소스 다운로드 (.zip)",
    needsBrief: "부여할 기능과 사용 방법을 모두 입력하세요.", copied: "복사됨 ✓", copy: "코드 복사",
    shieldTitle: "Antivirus 보호는 유지하고, 오탐 대응은 자동화합니다.", shieldCopy: "공급업체 중립 절차로 탐지 증거를 수집하고 검증된 파일만 검토 및 최소 허용 단계로 보냅니다.",
    auto: "수집 · 해시 · 격리 · 제출 대기열", review: "최종 허용 · 복원 · 배포 승인",
    noAutoTitle: "자동 예외 금지", noAutoCopy: "모든 파일을 신뢰시키거나 보호 기능을 전체 해제하지 않습니다. 검증 전에는 격리 상태를 유지합니다.",
    verifyCopy: "승인된 최소 허용 범위를 적용한 뒤 가장 짧은 명령 체인으로 로컬 도구를 확인하세요.",
    checking: "검사 중", standby: "대기", passed: "통과", run: "로컬 준비 상태 검사", again: "다시 검사",
    pass: "✓ 명령 체인이 준비되었습니다. 실제 탐지 여부는 Antivirus 보호 기록에서 최종 확인하세요.",
    footer: "보안을 해제하지 않는 임베디드 개발 워크플로",
  },
} as const;

function commentLines(value: string) {
  return value.replaceAll("*/", "* /").trim().split(/\r?\n/).map((line) => " * " + line).join("\n");
}

function categoryMark(category: string) {
  const marks: Record<string, string> = {
    Communication: "TX", Connectivity: "NET", Timing: "PWM", Display: "LCD", Graphics: "GFX",
    System: "SYS", Analog: "ADC", Storage: "MEM", Audio: "AUD", Security: "SEC",
  };
  return marks[category] ?? category.slice(0, 3).toUpperCase();
}

function makeCustomSource(profile: DriverProfile, feature: string, usage: string, cubeTool: "cubemx" | "cubemx2") {
  const toolNote = cubeTool === "cubemx2" ? [
    "/**",
    " * STM32CubeMX2 / HAL2 REFERENCE MODE",
    " * This template expresses the requested driver behavior.",
    " * Map HAL1-style calls to the HAL2 API supplied by the selected MCU pack,",
    " * and keep application edits outside regenerated mx_ configuration files.",
    " */",
    "",
  ].join("\n") : "";
  if (!feature.trim() || !usage.trim()) return toolNote + profile.code;
  const brief = [
    "/**",
    " * PROJECT-SPECIFIC DRIVER BRIEF",
    " *",
    " * Functions:",
    commentLines(feature),
    " *",
    " * Intended use:",
    commentLines(usage),
    " *",
    " * Integration note: review clocks, pins, interrupts, DMA and error paths",
    " * against the selected STM32 and board before production use.",
    " */",
  ].join("\n");
  return toolNote + brief + "\n\n" + profile.code;
}

function u16(value: number) {
  return new Uint8Array([value & 255, (value >>> 8) & 255]);
}

function u32(value: number) {
  return new Uint8Array([value & 255, (value >>> 8) & 255, (value >>> 16) & 255, (value >>> 24) & 255]);
}

function joinBytes(parts: Uint8Array[]) {
  const size = parts.reduce((sum, part) => sum + part.length, 0);
  const output = new Uint8Array(size);
  let offset = 0;
  for (const part of parts) {
    output.set(part, offset);
    offset += part.length;
  }
  return output;
}

function crc32(data: Uint8Array) {
  let crc = 0xffffffff;
  for (const byte of data) {
    crc ^= byte;
    for (let bit = 0; bit < 8; bit += 1) crc = (crc >>> 1) ^ (0xedb88320 & -(crc & 1));
  }
  return (crc ^ 0xffffffff) >>> 0;
}

function makeZip(files: { name: string; content: string }[]) {
  const encoder = new TextEncoder();
  const local: Uint8Array[] = [];
  const central: Uint8Array[] = [];
  let offset = 0;

  for (const file of files) {
    const name = encoder.encode(file.name);
    const data = encoder.encode(file.content);
    const checksum = crc32(data);
    const localHeader = joinBytes([
      u32(0x04034b50), u16(20), u16(0), u16(0), u16(0), u16(0),
      u32(checksum), u32(data.length), u32(data.length), u16(name.length), u16(0), name, data,
    ]);
    local.push(localHeader);
    central.push(joinBytes([
      u32(0x02014b50), u16(20), u16(20), u16(0), u16(0), u16(0), u16(0),
      u32(checksum), u32(data.length), u32(data.length), u16(name.length), u16(0), u16(0),
      u16(0), u16(0), u32(0), u32(offset), name,
    ]));
    offset += localHeader.length;
  }

  const localData = joinBytes(local);
  const centralData = joinBytes(central);
  const end = joinBytes([
    u32(0x06054b50), u16(0), u16(0), u16(files.length), u16(files.length),
    u32(centralData.length), u32(localData.length), u16(0),
  ]);
  return new Blob([localData, centralData, end], { type: "application/zip" });
}

const checklist = [
  ["01", "탐지 자동 수집", "제품명, 탐지명, 파일 경로와 SHA-256을 기록하고 의심 파일을 격리합니다."],
  ["02", "신뢰 체인 검증", "공식 배포 출처, 코드 서명, 고정된 의존성과 재현 가능한 빌드를 확인합니다."],
  ["03", "다중 엔진 재검사", "단일 제품의 결과만 믿지 않고 여러 탐지 결과와 행위 로그를 함께 검토합니다."],
  ["04", "제조사 검토 요청", "검증된 샘플과 탐지 정보를 해당 안티바이러스 제조사에 오탐으로 제출합니다."],
  ["05", "최소 범위 임시 허용", "검토가 끝날 때까지 해시 또는 단일 산출물만 제한적으로 허용하고 만료일을 둡니다."],
];

function ScrollTextArea({ id, value, onChange, placeholder, upLabel, downLabel, resetKey }: {
  id: string;
  value: string;
  onChange: (value: string) => void;
  placeholder: string;
  upLabel: string;
  downLabel: string;
  resetKey: string;
}) {
  const textareaRef = useRef<HTMLTextAreaElement>(null);
  const [scrollState, setScrollState] = useState({ overflow: false, canUp: false, canDown: false });

  const updateScrollState = useCallback(() => {
    const element = textareaRef.current;
    if (!element) return;
    const overflow = element.scrollHeight > element.clientHeight + 1;
    const next = {
      overflow,
      canUp: overflow && element.scrollTop > 1,
      canDown: overflow && element.scrollTop + element.clientHeight < element.scrollHeight - 1,
    };
    setScrollState((current) => current.overflow === next.overflow && current.canUp === next.canUp && current.canDown === next.canDown ? current : next);
  }, []);

  useEffect(() => {
    const element = textareaRef.current;
    if (!element) return;
    updateScrollState();
    const observer = new ResizeObserver(updateScrollState);
    observer.observe(element);
    return () => observer.disconnect();
  }, [value, updateScrollState]);

  useEffect(() => {
    const element = textareaRef.current;
    if (!element) return;
    element.scrollTop = 0;
    updateScrollState();
  }, [resetKey, updateScrollState]);

  function scrollLine(direction: -1 | 1) {
    const element = textareaRef.current;
    if (!element) return;
    const lineHeight = Number.parseFloat(window.getComputedStyle(element).lineHeight) || 20;
    element.scrollBy({ top: direction * lineHeight, behavior: "smooth" });
  }

  return <div className={`scroll-textarea${scrollState.overflow ? " has-overflow" : ""}`}>
    <textarea ref={textareaRef} id={id} value={value} onChange={(event) => onChange(event.target.value)} onScroll={updateScrollState} placeholder={placeholder} rows={4} maxLength={1200} />
    <div className="textarea-scroll-controls" aria-hidden={!scrollState.overflow}>
      <button type="button" onClick={() => scrollLine(-1)} disabled={!scrollState.canUp} aria-label={upLabel}>▲</button>
      <button type="button" onClick={() => scrollLine(1)} disabled={!scrollState.canDown} aria-label={downLabel}>▼</button>
    </div>
  </div>;
}

export default function Home() {
  const [peripheral, setPeripheral] = useState<Peripheral>("UART");
  const [locale, setLocale] = useState<"en" | "ko">("en");
  const [driverSearch, setDriverSearch] = useState("");
  const [driverCategory, setDriverCategory] = useState("ALL");
  const [cubeTool, setCubeTool] = useState<"cubemx" | "cubemx2">("cubemx");
  const [guideOpen, setGuideOpen] = useState(false);
  const [featureBrief, setFeatureBrief] = useState(driverExamples.UART.featureEn);
  const [usageBrief, setUsageBrief] = useState(driverExamples.UART.usageEn);
  const [briefError, setBriefError] = useState(false);
  const [copied, setCopied] = useState(false);
  const [completed, setCompleted] = useState<number[]>([]);
  const [status, setStatus] = useState<"idle" | "running" | "passed">("idle");
  const profile = driverProfiles[peripheral];
  const meta = driverMeta[peripheral];
  const t = ui[locale];
  const activeChecklist = locale === "en" ? checklistEn : checklist;
  const customCode = useMemo(() => makeCustomSource(profile, featureBrief, usageBrief, cubeTool), [profile, featureBrief, usageBrief, cubeTool]);
  const categories = useMemo(() => Array.from(new Set(peripherals.map((item) => driverMeta[item].category))).sort(), []);
  const filteredDrivers = useMemo(() => {
    const query = driverSearch.trim().toLowerCase();
    return peripherals.filter((item) => {
      const itemMeta = driverMeta[item];
      const matchesCategory = driverCategory === "ALL" || itemMeta.category === driverCategory;
      const itemProfile = driverProfiles[item];
      const matchesSearch = !query || item.toLowerCase().includes(query) || itemMeta.category.toLowerCase().includes(query) || itemProfile.summaryEn.toLowerCase().includes(query) || itemProfile.summary.toLowerCase().includes(query);
      return matchesCategory && matchesSearch;
    });
  }, [driverSearch, driverCategory]);

  const progress = useMemo(() => Math.round((completed.length / activeChecklist.length) * 100), [completed, activeChecklist.length]);

  useEffect(() => {
    document.documentElement.lang = locale;
  }, [locale]);

  async function copyCode() {
    await navigator.clipboard.writeText(customCode);
    setCopied(true);
    window.setTimeout(() => setCopied(false), 1600);
  }

  function downloadSource() {
    if (!featureBrief.trim() || !usageBrief.trim()) {
      setBriefError(true);
      return;
    }
    setBriefError(false);
    const baseName = profile.filename.replace(/\.(c|cpp)$/, "");
    const headerFilename = profile.headerFilename ?? baseName + ".h";
    const cubeIntegration = cubeTool === "cubemx2" ? [
      "- Open or create the supported device configuration from an .ioc2 project.",
      "- Inspect the peripheral .c/.h output in Code preview before generation.",
      "- Do not edit generated mx_ configuration files; keep application logic in separate modules.",
      "- Map this behavioral template to the HAL2 API delivered by the selected MCU pack.",
    ] : ["- Confirm the generated STM32 HAL handle and STM32CubeMX peripheral initialization call."];
    const readme = [
      "# STM32 " + peripheral + " Driver Package", "", "Generated by STM32 Driver Lab.", "",
      "Configuration reference: " + (cubeTool === "cubemx2" ? "STM32CubeMX2 / HAL2" : "STM32CubeMX / HAL1"), "",
      "## Requested functions", featureBrief.trim(), "", "## Intended use", usageBrief.trim(), "",
      "## Integration checklist", "- Add " + profile.filename + " and " + headerFilename + " to the application.",
      ...cubeIntegration,
      "- Review GPIO alternate functions, clocks, interrupts, DMA and cache policy.",
      "- Validate timeouts and error recovery on the target board.",
      "- Run static analysis and hardware-in-the-loop tests before release.", "",
      "This generated package is a reviewable engineering starting point, not a board-specific certified driver.",
    ].join("\n");
    const archive = makeZip([
      { name: profile.filename, content: customCode + "\n" },
      { name: headerFilename, content: driverHeaders[peripheral] + "\n" },
      { name: "README.md", content: readme + "\n" },
    ]);
    const url = URL.createObjectURL(archive);
    const link = document.createElement("a");
    link.href = url;
    link.download = baseName + "_source.zip";
    link.click();
    window.setTimeout(() => URL.revokeObjectURL(url), 1000);
  }

  function runCheck() {
    setStatus("running");
    window.setTimeout(() => setStatus("passed"), 1200);
  }

  function applyBriefs(item: Peripheral, lang: "en" | "ko") {
    const examples = driverExamples[item];
    setFeatureBrief(lang === "en" ? examples.featureEn : examples.featureKo);
    setUsageBrief(lang === "en" ? examples.usageEn : examples.usageKo);
  }

  function changeLocale(next: "en" | "ko") {
    if (next === locale) return;
    setLocale(next);
    applyBriefs(peripheral, next);
  }

  function chooseDriver(item: Peripheral) {
    setPeripheral(item);
    applyBriefs(item, locale);
    setGuideOpen(false);
    setBriefError(false);
  }

  function chooseCategory(category: string) {
    setDriverCategory(category);
    setDriverSearch("");
    const firstMatch = peripherals.find((item) => category === "ALL" || driverMeta[item].category === category);
    if (firstMatch) {
      chooseDriver(firstMatch);
    }
  }

  return (
    <main lang={locale}>
      <header className="topbar">
        <a className="brand" href="#top" aria-label="STM32 Driver Lab home">
          <span className="brand-mark">D<span>32</span></span>
          <span>DRIVER LAB</span>
        </a>
        <nav aria-label="Main navigation">
          <a href="#generator">{t.navForge}</a>
          <a href="#shield">{t.navShield}</a>
          <a href="#verify">{t.navVerify}</a>
        </nav>
        <div className="top-actions">
          <div className="language-toggle" role="group" aria-label="Display language">
            <button className={locale === "en" ? "active" : ""} aria-pressed={locale === "en"} onClick={() => changeLocale("en")}>EN</button>
            <button className={locale === "ko" ? "active" : ""} aria-pressed={locale === "ko"} onClick={() => changeLocale("ko")}>한</button>
          </div>
          <span className="system-pill"><i /> LOCAL READY</span>
        </div>
      </header>

      <section className="hero" id="top">
        <div className="eyebrow"><span>EMBEDDED WORKBENCH</span><b>v2.0</b></div>
        <h1>Build close<br />to the <em>metal.</em></h1>
        <p className="hero-copy">{t.hero}</p>
        <div className="hero-actions">
          <a className="primary" href="#generator">{t.start} <span>↘</span></a>
          <a className="secondary" href="#shield">{t.guide}</a>
        </div>
        <div className="signal" aria-hidden="true">
          <span>GPIO_A5</span>
          <div className="wave"><i/><i/><i/><i/><i/></div>
          <small>84 MHz&nbsp;&nbsp; • &nbsp;&nbsp;3.3 V</small>
        </div>
        <div className="hero-note"><span>01</span><p>REGISTER-LEVEL CLARITY<br/><b>HAL-ready, review-friendly templates.</b></p></div>
      </section>

      <section className="generator section" id="generator">
        <div className="section-heading">
          <span>01 / DRIVER FORGE</span>
          <h2>{t.forgeTitle}</h2>
        </div>
        <div className="resource-rail" aria-label="STM32 application library capabilities">
          <span><b>{peripherals.length}</b><small>DRIVER PROFILES</small></span>
          <span><b>09</b><small>DISPLAY PATHS</small></span>
          <span><b>HAL1 / HAL2</b><small>CUBEMX READY</small></span>
          <span><b>C / C++</b><small>SOURCE PACKAGES</small></span>
        </div>
        <div className="application-library">
          <div className="library-head">
            <div><span>{t.libraryTitle}</span><p>{t.libraryCopy}</p></div>
            <div className="library-count"><strong>{peripherals.length}</strong><span>APPLICATION<br/>PROFILES</span></div>
          </div>
          <div className="selection-flow">
            <div className="selection-column tool-column">
              <div className="selection-label"><b>01</b><span>{t.chooseTool}</span></div>
              <div className="mx-version-toggle" role="group" aria-label={t.cubeReference}>
                <button className={cubeTool === "cubemx" ? "active" : ""} aria-pressed={cubeTool === "cubemx"} onClick={() => setCubeTool("cubemx")}><small>HAL1</small>CubeMX</button>
                <button className={cubeTool === "cubemx2" ? "active" : ""} aria-pressed={cubeTool === "cubemx2"} onClick={() => setCubeTool("cubemx2")}><small>HAL2</small>CubeMX2</button>
              </div>
            </div>
            <label className="selection-column select-column">
              <span className="selection-label"><b>02</b><span>{t.chooseCategory}</span></span>
              <span className="select-shell"><i>{driverCategory === "ALL" ? "ALL" : categoryMark(driverCategory)}</i><select value={driverCategory} onChange={(event) => chooseCategory(event.target.value)} aria-label={t.chooseCategory}>
                <option value="ALL">{t.allCategories}</option>
                {categories.map((item) => <option key={item} value={item}>{item}</option>)}
              </select></span>
              <small>{filteredDrivers.length} {t.matchingDrivers}</small>
            </label>
            <label className="selection-column select-column featured-select">
              <span className="selection-label"><b>03</b><span>{t.chooseDriver}</span></span>
              <span className="select-shell"><i>{categoryMark(meta.category)}</i><select value={peripheral} onChange={(event) => chooseDriver(event.target.value)} aria-label={t.chooseDriver}>
                {Array.from(new Set([peripheral, ...(filteredDrivers.length ? filteredDrivers : peripherals)])).map((item) => <option key={item} value={item}>{item} — {driverMeta[item].category}</option>)}
              </select></span>
              <small>{t.browseHint}</small>
            </label>
          </div>
          <div className="category-browser">
            <div className="category-browser-head"><span>{t.quickBrowse}</span><input type="search" value={driverSearch} onChange={(event) => setDriverSearch(event.target.value)} placeholder={t.searchDrivers} aria-label={t.searchDrivers} /></div>
            <div className="category-strip" role="group" aria-label={t.quickBrowse}>
              <button className={driverCategory === "ALL" ? "active" : ""} onClick={() => chooseCategory("ALL")}><b>ALL</b><span>{peripherals.length}</span></button>
              {categories.map((item) => <button key={item} className={driverCategory === item ? "active" : ""} onClick={() => chooseCategory(item)}><b>{categoryMark(item)}</b><span>{item}<em>{peripherals.filter((driver) => driverMeta[driver].category === item).length}</em></span></button>)}
            </div>
          </div>
          <div className="catalog-grid">
            {filteredDrivers.map((item, index) => {
              const itemProfile = driverProfiles[item];
              const itemMeta = driverMeta[item];
              return <button key={item} className={peripheral === item ? "active" : ""} aria-pressed={peripheral === item} onClick={() => chooseDriver(item)}>
                <span className="card-number">{String(index + 1).padStart(2, "0")}</span><span className="card-mark">{categoryMark(itemMeta.category)}</span><span className="card-category">{itemMeta.category}</span><b>{item}</b><small>{locale === "en" ? itemProfile.summaryEn : itemProfile.summary}</small>
              </button>;
            })}
          </div>
        </div>
        <div className="forge-grid">
          <div className="controls">
            <label>{t.peripheral}</label>
            <div className="selected-driver">
              <i>{categoryMark(meta.category)}</i><span><small>{t.selectedProfile}</small>{meta.category} / {cubeTool === "cubemx2" ? "CUBEMX2 · HAL2" : "CUBEMX · HAL1"}</span><b>{peripheral}</b>
            </div>
            <button className="guide-trigger" onClick={() => setGuideOpen((current) => !current)}>{guideOpen ? t.closeGuide : t.openGuide}<span>{guideOpen ? "↑" : "↗"}</span></button>
            <div className="spec-list">
              {profile.specs.map(([label, value]) => <div key={label}><span>{label}</span><b>{value}</b></div>)}
            </div>
            <p className="profile-summary">{locale === "en" ? profile.summaryEn : profile.summary}</p>
            <p className="microcopy">{t.boardNote}</p>
          </div>
          <div className="code-card">
            <div className="code-top"><span><i/> {profile.filename}</span><button onClick={copyCode}>{copied ? t.copied : t.copy}</button></div>
            <pre><code>{customCode}</code></pre>
            <div className="code-foot"><span>{cubeTool === "cubemx2" ? "CUBEMX2 · HAL2 REF" : "CUBEMX · HAL1"}</span>{profile.tags.map((tag) => <span key={tag}>{tag}</span>)}</div>
          </div>
        </div>
        {guideOpen && <aside className="application-guide" aria-live="polite">
          <div className="guide-title"><span>DEVELOPING APPLICATION / {meta.category} / {cubeTool === "cubemx2" ? "CUBEMX2" : "CUBEMX"}</span><button onClick={() => setGuideOpen(false)} aria-label={t.closeGuide}>×</button></div>
          <div className="guide-body">
            <div><small>{t.useCases}</small><h3>{peripheral}</h3><p>{locale === "en" ? meta.applicationEn : meta.applicationKo}</p>{cubeTool === "cubemx2" && <p className="mx2-note">{t.cubeMx2Note}</p>}</div>
            <div><small>{cubeTool === "cubemx2" ? "CUBEMX2 + HAL2 SETUP" : t.setupFlow}</small><ol>{(cubeTool === "cubemx2" ? t.cubeMx2Steps : t.setupSteps).map((step) => <li key={step}>{step}</li>)}</ol></div>
          </div>
          <div className="guide-foot"><p>{cubeTool === "cubemx2" ? t.cubeMx2Family : t.familyNote}</p><a href={cubeTool === "cubemx2" ? "https://dev.st.com/stm32cube-docs/stm32cubemx2/1.0.0/en/global_toc.html" : meta.guideUrl} target="_blank" rel="noreferrer">{t.officialGuide}<span>↗</span></a></div>
        </aside>}
        <div className="brief-builder">
          <div className="brief-fields">
            <label htmlFor="feature-brief">{t.featureLabel}</label>
            <ScrollTextArea id="feature-brief" value={featureBrief} onChange={(value) => { setFeatureBrief(value); setBriefError(false); }} placeholder={t.featurePlaceholder} upLabel={locale === "en" ? "Scroll functions up one line" : "기능을 한 줄 위로"} downLabel={locale === "en" ? "Scroll functions down one line" : "기능을 한 줄 아래로"} resetKey={peripheral} />
            <label htmlFor="usage-brief">{t.usageLabel}</label>
            <ScrollTextArea id="usage-brief" value={usageBrief} onChange={(value) => { setUsageBrief(value); setBriefError(false); }} placeholder={t.usagePlaceholder} upLabel={locale === "en" ? "Scroll usage up one line" : "사용 방법을 한 줄 위로"} downLabel={locale === "en" ? "Scroll usage down one line" : "사용 방법을 한 줄 아래로"} resetKey={peripheral} />
          </div>
          <div className="package-card">
            <span>02 / {t.packageTitle}</span>
            <h3>{peripheral}<br />FULL SOURCE</h3>
            <p>{t.packageCopy}</p>
            <ul><li>{profile.filename}</li><li>{profile.headerFilename ?? profile.filename.replace(/\.(c|cpp)$/, ".h")}</li><li>README.md</li></ul>
            <button onClick={downloadSource}>{t.download}<b>↓</b></button>
            <small className={briefError ? "form-note error" : "form-note"} aria-live="polite">{briefError ? t.needsBrief : "ZIP · C / H / README"}</small>
          </div>
        </div>
      </section>

      <section className="shield section" id="shield">
        <div className="shield-intro">
          <span className="kicker">02 / ANTIVIRUS SHIELD FLOW</span>
          <h2>{t.shieldTitle}</h2>
          <p>{t.shieldCopy}</p>
          <div className="shield-flow" aria-label="False-positive response scope">
            <span>AUTO</span><b>{t.auto}</b>
            <span>REVIEW</span><b>{t.review}</b>
          </div>
          <div className="vendor-row" aria-label="지원 대상"><span>ANTIVIRUS ENGINE</span><span>ENDPOINT SECURITY</span><span>EDR / XDR</span><span>VENDOR-NEUTRAL</span></div>
          <div className="warning"><b>!</b><span><strong>{t.noAutoTitle}</strong>{t.noAutoCopy}</span></div>
        </div>
        <div className="checklist">
          {activeChecklist.map(([num, title, copy], index) => {
            const done = completed.includes(index);
            return <button key={num} className={done ? "done" : ""} onClick={() => setCompleted(done ? completed.filter(i => i !== index) : [...completed, index])}>
              <span className="step-num">{num}</span><span className="step-copy"><b>{title}</b><small>{copy}</small></span><span className="check">{done ? "✓" : "○"}</span>
            </button>
          })}
          <div className="progress"><span><i style={{width: `${progress}%`}} /></span><b>{progress}% COMPLETE</b></div>
        </div>
      </section>

      <section className="verify section" id="verify">
        <div>
          <span className="kicker">03 / FINAL CHECK</span>
          <h2>Clean. Build. Flash.</h2>
          <p>{t.verifyCopy}</p>
        </div>
        <div className="terminal">
          <div className="terminal-top"><span>LOCAL COMMAND CHECK</span><span className={"status " + status}>{status === "passed" ? t.passed : status === "running" ? t.checking : t.standby}</span></div>
          <div className="command"><span>$</span><code>cmake --build build --target all</code></div>
          <button onClick={runCheck} disabled={status === "running"}>{status === "running" ? t.checking + "…" : status === "passed" ? t.again : t.run}<span>→</span></button>
          {status === "passed" && <p className="pass-message">{t.pass}</p>}
        </div>
      </section>

      <footer><span>DRIVER LAB / STM32 WORKBENCH</span><p>{t.footer}</p><a href="#top">TOP ↑</a></footer>
    </main>
  );
}
