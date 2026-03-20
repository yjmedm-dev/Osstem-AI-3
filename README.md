# Osstem AI - 해외법인 재무제표 검증 및 분석 시스템

> 해외법인이 제출한 재무제표 작성 자료의 적합성을 자동으로 검증하고 분석하는 AI 기반 솔루션

---

## 프로젝트 개요

해외 각 법인에서 제출하는 재무제표 작성 자료(Trial Balance, 계정별 명세서 등)를 수집하여 **데이터 정합성 검증**, **회계기준 적합성 분석**, **이상 항목 탐지** 등을 수행합니다. 수작업으로 이루어지던 검토 프로세스를 자동화하여 결산 효율성을 높이고 오류 리스크를 줄이는 것을 목표로 합니다.

---

## 주요 기능

### 1. 데이터 수집 및 전처리
- 해외법인별 재무 데이터 파일 수집 (Excel, CSV 등)
- 통화 코드, 계정 코드, 법인 코드 등 기준 정보 매핑
- 환율 적용 및 원화 환산 처리

### 2. 적합성 검증 (Validation)
- **계정 코드 유효성** 검사 (본사 계정 체계 대조)
- **대차 균형 검증** (차변 합계 = 대변 합계)
- **필수 항목 누락 여부** 확인
- **기간 비교 이상치** 탐지 (전기 대비 급격한 변동 항목)
- **그룹 내부거래 정합성** 검증 (상계 대상 거래 확인)

### 3. 재무 분석 (Analysis)
- 법인별 손익 현황 요약 및 시각화
- 주요 재무비율 산출 (유동비율, 부채비율, 영업이익률 등)
- 전기 대비 증감 분석 및 원인 분류
- 예산 대비 실적 비교 분석

### 4. 결과 리포트 생성
- 법인별 검증 결과 요약 리포트 자동 생성
- 오류/경고 항목 목록 및 조치 가이드 제공
- 경영진 보고용 대시보드 데이터 출력

---

## 기술 스택

| 구분 | 기술 |
|------|------|
| Language | Python 3.10+ |
| AI / LLM | Claude API (Anthropic) |
| 데이터 처리 | Pandas, OpenPyXL |
| 시각화 | Matplotlib, Plotly |
| 보고서 | Jinja2, ReportLab |
| 인프라 | (추후 작성) |

---

## 디렉토리 구조

```
Osstem-AI-3/
├── data/
│   ├── input/          # 해외법인 제출 원본 파일
│   ├── processed/      # 전처리 완료 데이터
│   └── reference/      # 계정코드 등 기준 데이터
├── src/
│   ├── ingestion/      # 데이터 수집 및 파싱
│   ├── validation/     # 적합성 검증 로직
│   ├── analysis/       # 재무 분석 모듈
│   └── report/         # 리포트 생성
├── tests/              # 단위 테스트
├── notebooks/          # 분석 탐색용 Jupyter Notebook
├── config/             # 설정 파일 (법인 목록, 계정 매핑 등)
└── README.md
```

---

## 시작하기

### 사전 요구사항

- Python 3.10 이상
- pip 패키지 관리자
- Anthropic API Key (Claude 사용 시)

### 설치

```bash
git clone https://github.com/yjmedm-dev/Osstem-AI-3.git
cd Osstem-AI-3
pip install -r requirements.txt
```

### 환경 변수 설정

```bash
cp .env.example .env
# .env 파일에 API 키 및 설정 입력
```

```
ANTHROPIC_API_KEY=your_api_key_here
```

### 실행

```bash
python src/main.py --corp all --period 2025Q4
```

---

## 검증 규칙 예시

| 규칙 ID | 항목 | 유형 | 설명 |
|---------|------|------|------|
| V001 | 대차 균형 | Error | 차변/대변 합계 불일치 시 오류 |
| V002 | 계정 코드 | Error | 본사 계정 체계 미등록 코드 |
| V003 | 필수 계정 | Warning | 법인 유형별 필수 계정 누락 |
| V004 | 이상 변동 | Warning | 전기 대비 ±50% 초과 변동 항목 |
| V005 | 내부거래 | Error | 상계 대상 내부거래 금액 불일치 |

---

## 기여 방법

1. 이 레포지토리를 Fork합니다.
2. 새 브랜치를 생성합니다. (`git checkout -b feature/기능명`)
3. 변경사항을 커밋합니다. (`git commit -m 'feat: 기능 설명'`)
4. 브랜치에 Push합니다. (`git push origin feature/기능명`)
5. Pull Request를 생성합니다.

---

## 라이선스

이 프로젝트는 내부 업무용으로 작성되었습니다. 외부 배포 시 별도 협의가 필요합니다.

---

*Osstem Implant Co., Ltd. — Finance & Digital Innovation Team*
