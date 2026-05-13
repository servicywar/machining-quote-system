# 가공 견적 AI 엔진 PRO v2.1

협력사 단가를 역산하여 재료비·가공비를 분리하고 적정성을 판정하는 설치형 견적 감사 시스템입니다.

## 빠른 시작

```bash
# 1. 패키지 설치
pip install -r requirements.txt

# 2. 실행
streamlit run main.py
```

브라우저에서 http://localhost:8501 접속

## STEP 파일 분석 (선택)

```bash
conda install -c conda-forge pythonocc-core
```

설치 없이도 수동 입력 모드로 모든 역산 기능을 사용할 수 있습니다.

## 역산 공식

```
재료비    = 중량(kg) × 시세(원/kg) × (1 + 로스율)
실질가공비 = 협력사단가 - 재료비 - 후처리비
적정가공비 = 임률(원/h) × 가공시간(h) × 난이도계수
오차율(%) = 실질가공비 / 적정가공비 × 100
```

## 판정 기준

| 오차율 | 판정 | 권장 액션 |
|--------|------|-----------|
| > 120% | 경고·과다 | 항목별 소명 요청 |
| 110~120% | 주의 | 단가 재협의 |
| 90~110% | 신뢰 | 발주 진행 |
| 80~90% | 관찰 | 규격 확인 |
| < 80% | 경고·저가 | 품질 조건 명기 |

## 파일 구조

```
machining_quote_system/
├── main.py                  # Streamlit UI
├── engine/
│   ├── auditor.py           # 역산 판정 (핵심)
│   ├── material.py          # 재료비 계산
│   ├── difficulty.py        # 난이도 등급
│   └── analyzer.py          # STEP 분석 (pythonOCC)
├── data/
│   ├── market_prices.xlsx   # 시세 오버라이드
│   ├── material_master.xlsx # 소재 마스터
│   └── history_log.db       # 감사 이력 (자동 생성)
└── requirements.txt
```
