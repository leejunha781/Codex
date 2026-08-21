# Taeha Dispenser Control Simulation

이 프로젝트는 MATLAB 없이도 태하 디스펜서의 상위 공정 제어를 학습하고 회귀 시험할 수 있는 실행 가능한 기준 모델입니다. 실제 모터 전류 루프나 유체 CFD를 대체하지 않습니다.

포함 범위:

- 재료·노즐·온도별 feed-forward
- 센서 지연과 잔류 토출(dribble)을 포함한 shot 모델
- 신뢰도·프로파일 격리·보정률 제한을 적용한 shot-to-shot 보정
- A/B 독립 계측 기반 혼합비 감시
- servo, stepper, brushed DC, BLDC/AC actuator 응답 추상화
- CRC32 및 이중 슬롯 보정 데이터 복구
- 결정론적 seed를 사용한 반복 shot 결과와 PNG/CSV/JSON 증거

실행:

```powershell
python -m pip install -e .
python -m taeha_dispense_sim.runner --shots 40 --output outputs/baseline
python -m unittest discover -s tests -v
```

설치 없이도 다음처럼 실행할 수 있습니다.

```powershell
$env:PYTHONPATH = "$PWD\src"
python -m taeha_dispense_sim.runner --shots 40 --output outputs/baseline
```

결과는 제어 알고리즘 비교용 `SIMULATED` 증거입니다. 실제 중량 정확도, A/B 비율, 안전 정지, ADC 지연, 모터 토크 및 HIL 통과는 정확한 보드·기구·재료·노즐·센서 조합에서 별도로 입증해야 합니다.
