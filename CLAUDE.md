# CLAUDE.md — Osstem AI 프로젝트 컨텍스트

## 프로젝트 목적

해외법인 10곳에서 매월 제출하는 재무 엑셀 파일을 자동으로 수집·검증·분석하여
담당자의 수작업 부담을 줄이고 결산 오류를 사전에 방지하는 시스템.

담당 부서: **Osstem Implant — 해외관리2팀**
핵심 업무 3가지: 법인 월마감 / 계획대비 차질사유 분석 / 롤링계획 관리

---

## 기술 스택

| 역할 | 기술 |
|------|------|
| 언어 | Python 3.11+ |
| 데이터 처리 | Pandas, OpenPyXL |
| AI 분석 | Claude API (`claude-sonnet-4-6`) |
| DB | SQLite(개발) / PostgreSQL(운영) |
| ORM | SQLAlchemy |
| 보고서 | Jinja2, WeasyPrint, XlsxWriter |
| CLI | Click, Rich |
| 이상값 탐지 | SciPy, Scikit-learn |

---

## 디렉토리 구조

```
Osstem-AI-3/
├── main.py                        # 진입점
├── config/
│   ├── settings.py                # 환경 변수, API 키
│   ├── accounts_master.yaml       # 계정과목 마스터
│   └── validation_rules.yaml      # 검증 규칙 정의
├── models/                        # 데이터 모델 (Pydantic or dataclass)
│   ├── subsidiary.py
│   ├── account.py
│   ├── trial_balance.py
│   └── validation_result.py
├── ingestion/                     # 파일 파싱 및 정규화
│   ├── excel_parser.py
│   ├── schema_mapper.py
│   └── data_normalizer.py
├── validation/                    # 검증 엔진
│   ├── engine.py
│   └── rules/
│       ├── base_rule.py
│       ├── arithmetic_rules.py
│       ├── accounting_rules.py
│       ├── period_rules.py
│       ├── anomaly_rules.py
│       └── provision_rules.py     # 충당금 전용 검증
├── analysis/                      # Claude API 연동
│   ├── claude_client.py
│   ├── prompt_builder.py
│   └── recommendation_engine.py
├── reporting/
│   ├── excel_reporter.py
│   ├── pdf_reporter.py
│   └── templates/
├── utils/
│   ├── currency.py
│   ├── date_utils.py
│   └── exceptions.py
├── data/
│   ├── input/                     # 법인 제출 원본 엑셀 (절대 편집 금지)
│   ├── processed/
│   └── reference/                 # 계정코드·환율 등 기준 데이터
└── tests/
```

---

## 도메인 지식 — 충당금 8종

해외법인은 아래 8종의 충당금을 운용한다. 코드 작성 시 이 목록을 기준으로 삼는다.

| # | 충당금 | 설정 기준 | 비고 |
|---|--------|-----------|------|
| 1 | 재고자산평가충당금 | 저속·불용·시가하락 재고 × 요율 | |
| 2 | 대손충당금 | 매출채권 Aging × 구간별 요율 | |
| 3 | 퇴직급여충당금 | 기말 퇴직금 추계 − 사외적립 | **의무 있는 법인만** 적용 |
| 4 | 리스회계충당금 | ROU자산·리스부채 (IFRS 16) | |
| 5 | 반품충당금 | 과거 반품률 × 당기 매출 | |
| 6 | 단품충당금 | 단종·사양화 SKU 재고 × 별도 요율 | |
| 7 | FOC충당금 | 무상공급 예상 수량 × 원가 | FOC = Free of Charge |
| 8 | 수익인식충당금 | 조건 미충족 계약 잔액 (IFRS 15) | |

---

## 도메인 지식 — 검증 규칙 ID 체계

| Prefix | 분류 |
|--------|------|
| `AR-`  | 산술 검증 (Arithmetic) |
| `PR-`  | 기간 비교 (Period) |
| `AN-`  | 이상값 탐지 (Anomaly) |
| `PV-`  | 충당금 검증 (Provision) |
| `IT-`  | 내부거래 검증 (Intercompany) |

현재 정의된 규칙:

| 규칙 ID | 내용 | 심각도 |
|---------|------|--------|
| AR-001 | 차변 합계 = 대변 합계 | CRITICAL |
| AR-002 | 자산 = 부채 + 자본 | CRITICAL |
| AR-003 | 이익잉여금 검증 | ERROR |
| PR-001 | 전기 대비 30% 초과 변동 | WARNING |
| AN-001 | 필수 계정 잔액 0 탐지 | WARNING |
| AN-002 | 자산 계정 음수 잔액 | ERROR |
| AN-003 | 비정상적 반올림 집중 | WARNING |

---

## 코딩 컨벤션

- 검증 규칙은 `validation/rules/base_rule.py`의 `BaseRule`을 상속해서 구현한다.
- 충당금 관련 규칙은 반드시 `validation/rules/provision_rules.py`에 모은다.
- 법인 식별자는 `subsidiary_code` (문자열, 예: `"KR01"`, `"US02"`) 를 표준으로 쓴다.
- 금액은 내부적으로 **원화(KRW) 기준 float**로 통일한다. 외화 원본은 별도 필드에 보존한다.
- 환율은 `utils/currency.py`의 `convert()` 함수를 통해서만 변환한다. 직접 곱셈 금지.
- 파일 경로는 하드코딩 금지. `config/settings.py`의 경로 상수를 사용한다.
- `data/input/` 안의 원본 파일은 절대 수정하지 않는다. 읽기 전용으로 취급한다.

---

## 주요 업무 흐름

### 월마감
```
법인 파일 수신 확인 → 형식 검증 → 산술/계정 검증 → 충당금 검증
→ 내부거래 상계 검증 → 환율 환산 → 연결 Trial Balance 집계 → 확정
```

### 차질사유 분석
```
실적 집계 → 계획 대조 → 차이율 산출 → 임계치 플래그
→ Claude API 원인 초안 생성 → 담당자 검토 → 보고서 출력
```

### 롤링계획
```
누계 실적 확인 → 잔여 기간 필요 이익 역산 → 민감도 분석
→ Base/Best/Worst 시나리오 계산 → 법인 피드백 반영 → 확정
```

---

## 주의사항

- 환경 변수(`ANTHROPIC_API_KEY` 등)는 `.env` 파일에서 읽는다. 코드에 직접 쓰지 않는다.
- 테스트용 샘플 데이터는 실제 법인 수치를 사용하지 않는다. `tests/fixtures/` 안에 익명화된 데이터를 사용한다.
- Claude API 호출 시 재무 수치가 포함된 프롬프트는 최소한으로 구성한다 (개인정보·영업비밀 보호).

---

## 구현 현황 (2026-04-23 기준)

### 참조 데이터 (소스 DB)

| 파일 | 설명 |
|------|------|
| `BSPL검토_우즈베키스탄_2603.xlsx` | UZ01 2603 BS/PL 검토용 소스 (시트: 월비교, 누계비교, BS, PL) |
| `BSPL검토_러시아_2603.xlsx` | RU01 2603 BS/PL 검토용 소스 |
| `BSPL검토_우크라이나_2603.xlsx` | UA01 2603 BS/PL 검토용 소스 |
| `BSPL검토_카자흐스탄_2603.xlsx` | KZ01 2603 BS/PL 검토용 소스 |
| `우즈벡 마감자료_2602.xlsx` | UZ01 2602 마감 완성본 (시각화 목표 형식) |

### 완료된 파일

| 파일 | 역할 |
|------|------|
| `config/settings.py` | 경로 상수, API 키, 법인 코드 목록, 충당금 의무 법인 목록 |
| `config/accounts_master.yaml` | 계정과목 마스터 (충당금 8종 포함) |
| `config/validation_rules.yaml` | 검증 규칙 정의 (AR/PR/AN/PV/IT 체계) |
| `models/subsidiary.py` | 법인 데이터 모델 |
| `models/account.py` | 계정과목 데이터 모델 |
| `models/trial_balance.py` | 시산표 행/전체 모델 |
| `models/validation_result.py` | 검증 결과·이슈·심각도 모델 |
| `utils/currency.py` | 외화 → KRW 환산 (`convert()` 함수만 사용) |
| `utils/date_utils.py` | 기간 문자열 유틸리티 |
| `utils/exceptions.py` | 프로젝트 전용 예외 클래스 |
| `ingestion/excel_parser.py` | 엑셀 파일 파싱 → TrialBalance 변환 |
| `ingestion/schema_mapper.py` | 법인별 컬럼명 → 표준 컬럼명 매핑 |
| `ingestion/data_normalizer.py` | 외화 금액 원화 환산 정규화 |
| `validation/rules/base_rule.py` | 검증 규칙 공통 기반 클래스 |
| `validation/rules/arithmetic_rules.py` | AR-001~003 산술 검증 |
| `validation/rules/accounting_rules.py` | AC-001 계정 코드 유효성 검사 |
| `validation/rules/period_rules.py` | PR-001 전기 대비 변동 검증 |
| `validation/rules/anomaly_rules.py` | AN-001~003 이상값 탐지 |
| `validation/rules/provision_rules.py` | PV-001~009 충당금 8종 + 환입 검증 |
| `validation/engine.py` | 모든 규칙 일괄 실행 엔진 |
| `main.py` | CLI 진입점 (`python main.py validate --period 2025-03`) |

### 미구현 (다음 단계)
- `analysis/` — Claude API 연동 (차질사유 초안 생성)
- `reporting/` — 엑셀·PDF 보고서 출력
- `tests/` — 단위 테스트
