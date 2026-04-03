from pathlib import Path
from dotenv import load_dotenv
import os

load_dotenv()

# 프로젝트 루트
BASE_DIR = Path(__file__).resolve().parent.parent

# 데이터 경로
DATA_INPUT_DIR   = BASE_DIR / os.getenv("DATA_INPUT_DIR",     "data/input")
DATA_PROCESSED_DIR = BASE_DIR / os.getenv("DATA_PROCESSED_DIR", "data/processed")
DATA_REFERENCE_DIR = BASE_DIR / os.getenv("DATA_REFERENCE_DIR", "data/reference")

# 설정 파일 경로
ACCOUNTS_MASTER_PATH    = BASE_DIR / "config" / "accounts_master.yaml"
VALIDATION_RULES_PATH   = BASE_DIR / "config" / "validation_rules.yaml"

# AI
ANTHROPIC_API_KEY = os.getenv("ANTHROPIC_API_KEY", "")
CLAUDE_MODEL      = "claude-sonnet-4-6"

# 검증 임계치
PERIOD_CHANGE_THRESHOLD = 0.30   # 전기 대비 30% 초과 변동 → WARNING

# 법인 코드 목록 (추후 DB 또는 YAML로 이관 가능)
SUBSIDIARY_CODES = [
    "CN01", "CN02", "US01", "DE01", "AU01",
    "BR01", "IN01", "RU01", "VN01", "ID01",
]

# 퇴직급여충당금 의무 적용 법인 (국가별 법령 기준)
RETIREMENT_MANDATORY_SUBSIDIARIES = ["KR01", "JP01", "IN01"]
