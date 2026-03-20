# 프로젝트 개발 계획서 (PLAN)

> 해외법인 재무제표 적합성 검증 및 분석 시스템

---

## 전체 구조 한눈에 보기

```
엑셀 파일 입력
    ↓
데이터 읽기 & 정리 (Ingestion)
    ↓
오류 검증 (Validation Engine)
    ↓
AI 분석 (Claude API)
    ↓
결과 보고서 출력 (Excel / PDF)
```

---

## 폴더 및 파일 구조

```
Osstem-AI-3/
│
├── main.py                        # 프로그램 시작점
│
├── config/
│   ├── settings.py                # API 키, DB 주소 등 환경 설정
│   ├── accounts_master.yaml       # 계정과목 마스터 목록
│   └── validation_rules.yaml      # 검증 규칙 정의
│
├── models/                        # 데이터 구조 정의
│   ├── subsidiary.py              # 법인 정보
│   ├── account.py                 # 계정과목 정보
│   ├── trial_balance.py           # 시산표 데이터
│   └── validation_result.py       # 검증 결과
│
├── ingestion/                     # 파일 읽기 & 정리
│   ├── excel_parser.py            # 엑셀 파일 파싱
│   ├── csv_parser.py              # CSV 파일 파싱
│   ├── schema_mapper.py           # 법인별 계정코드 매핑
│   └── data_normalizer.py         # 통화·날짜 형식 정규화
│
├── validation/                    # 검증 엔진
│   ├── engine.py                  # 검증 총괄 실행기
│   └── rules/
│       ├── base_rule.py           # 규칙 공통 틀
│       ├── arithmetic_rules.py    # 산술 검증 (차변=대변 등)
│       ├── accounting_rules.py    # 회계 원칙 검증
│       ├── period_rules.py        # 전기 대비 비교 검증
│       └── anomaly_rules.py       # 이상값 탐지
│
├── analysis/                      # AI 분석
│   ├── claude_client.py           # Claude API 연결
│   ├── prompt_builder.py          # AI에게 보낼 질문 구성
│   └── recommendation_engine.py   # 분석 결과 정리
│
├── reporting/                     # 보고서 생성
│   ├── excel_reporter.py          # 엑셀 보고서
│   ├── pdf_reporter.py            # PDF 보고서
│   └── templates/                 # 보고서 양식
│
├── utils/                         # 공통 도구
│   ├── currency.py                # 환율 변환
│   ├── date_utils.py              # 날짜 계산
│   └── exceptions.py              # 오류 메시지 정의
│
├── tests/                         # 테스트 코드
├── requirements.txt               # 필요한 라이브러리 목록
├── .env                           # API 키 등 비밀 설정
├── PLAN.md                        # 이 파일 (개발 계획)
└── README.md                      # 프로젝트 소개
```

---

## 단계별 개발 로드맵

### Phase 1 — 기반 구축 (4주)
- [ ] 프로젝트 폴더 구조 생성
- [ ] 환경 설정 파일 (.env, requirements.txt)
- [ ] 데이터 모델 설계 (법인, 계정, 시산표)
- [ ] 계정과목 마스터 데이터 구축

### Phase 2 — 데이터 수집 모듈 (2주)
- [ ] 엑셀/CSV 파일 파싱 구현
- [ ] 법인별 계정코드 매핑 로직
- [ ] 통화·날짜 정규화 처리

### Phase 3 — 검증 엔진 (3주)
- [ ] 산술 검증 규칙 (차변=대변, 회계등식)
- [ ] 기간 비교 검증 (전기 대비 변동)
- [ ] 이상값 탐지 규칙
- [ ] 검증 결과 리포트 생성

### Phase 4 — AI 분석 연동 (2주)
- [ ] Claude API 클라이언트 구현
- [ ] 프롬프트 설계 및 최적화
- [ ] AI 응답 파싱 및 결과 저장

### Phase 5 — 보고서 출력 (2주)
- [ ] 엑셀 검증 보고서 생성
- [ ] PDF 요약 보고서 생성
- [ ] 이메일 알림 기능

### Phase 6 — 고도화 (추후)
- [ ] 환율 API 자동 연동
- [ ] 연결 결산 내부거래 자동 매칭
- [ ] 웹 대시보드 (선택)
- [ ] ERP 시스템 연동 (SAP 등)

---

## 주요 검증 규칙 목록

| 규칙 ID | 분류 | 내용 | 심각도 |
|---------|------|------|--------|
| AR-001 | 산술 | 차변 합계 = 대변 합계 | CRITICAL |
| AR-002 | 산술 | 자산 = 부채 + 자본 | CRITICAL |
| AR-003 | 산술 | 이익잉여금 검증 | ERROR |
| PR-001 | 기간 비교 | 전기 대비 30% 초과 변동 | WARNING |
| AN-001 | 이상값 | 필수 계정 잔액 0 탐지 | WARNING |
| AN-002 | 이상값 | 자산 계정 음수 잔액 | ERROR |
| AN-003 | 이상값 | 비정상적 반올림 집중 | WARNING |

---

## 사용 기술 스택

| 역할 | 기술 |
|------|------|
| 언어 | Python 3.11+ |
| AI 분석 | Claude API (Anthropic) |
| 데이터 처리 | Pandas, OpenPyXL |
| 데이터베이스 | SQLite (개발) / PostgreSQL (운영) |
| ORM | SQLAlchemy |
| 보고서 | Jinja2, WeasyPrint, XlsxWriter |
| CLI | Click, Rich |
| 이상값 탐지 | SciPy, Scikit-learn |
