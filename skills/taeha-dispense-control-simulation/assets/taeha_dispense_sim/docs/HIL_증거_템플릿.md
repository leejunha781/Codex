# 태하 디스펜서 HIL 증거 템플릿

## 식별

| 항목 | 값 |
|---|---|
| 요구사항 ID | |
| 시험 ID / revision | |
| 날짜 / 시험자 / 검토자 | |
| STM32 MCU / board revision | |
| fixture revision | |
| probe serial | |
| firmware ELF/BIN SHA-256 | |
| Git branch/revision | |
| toolchain/HAL/RTOS version | |
| material A/B and lot | |
| nozzle/mixer/pump/gear | |
| motor/drive and firmware | |
| profile/calibration schema and sequence | |
| instruments and calibration due date | |
| temperature/humidity/supply pressure | |

## 안전 사전조건

- [ ] motor/valve outputs disabled at reset
- [ ] polarity, electrical level, current/pressure/travel limits confirmed
- [ ] bench supply current limit configured
- [ ] emergency stop and containment checked
- [ ] fault restoration procedure reviewed
- [ ] correct target/probe/image verified
- [ ] watchdog and brownout policy enabled

## 시험 자극과 합격 기준

| 요구사항 | 자극 | 계측 | 합격 기준 | 결과 | 증거 링크 |
|---|---|---|---|---|---|
| | | | | | |

## 필수 fault injection

- [ ] A/B communication loss and stale command
- [ ] CRC and malformed frame
- [ ] load-cell open/short/saturation/freeze
- [ ] ADC/DMA stale or overrun
- [ ] motor slip/missed step/drive fault
- [ ] valve stuck or delayed
- [ ] empty supply, air/cavitation, blockage
- [ ] pressure/temperature limit
- [ ] ratio out of tolerance
- [ ] queue overflow/task stall/watchdog
- [ ] reset/brownout at each calibration write phase
- [ ] corrupted latest slot and both-slot invalid
- [ ] safe state while interlock is asserted at reset

## 반복 shot 통계

| 지표 | A | B | Total/Ratio |
|---|---:|---:|---:|
| sample count | | | |
| rejected count | | | |
| mean error | | | |
| standard deviation | | | |
| 95% confidence interval | | | |
| min/max | | | |
| constraint violations | | | |

## 원시 증거

- [ ] timestamped raw ADC
- [ ] filtered mass and stability state
- [ ] motor command/current/speed/position
- [ ] pressure, temperature, valve state
- [ ] state machine and fault bits
- [ ] logic analyzer/scope capture
- [ ] calibration/profile record readback
- [ ] power-loss point and recovery log
- [ ] firmware and test artifact manifest

## 판정

- Evidence class: HIL / PHYSICAL
- Pass / fail / blocked:
- Residual risks:
- Hardware-only claims still UNVERIFIED:
- Required corrective action:
- Reviewer disposition:
