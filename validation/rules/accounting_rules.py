from pathlib import Path
import yaml

from models.trial_balance import TrialBalance
from models.validation_result import ValidationResult, ValidationIssue, Severity
from validation.rules.base_rule import BaseRule

_YAML_PATH = Path(__file__).parent.parent.parent / "config" / "accounts_master.yaml"


def _load_account_codes() -> set[str]:
    """accounts_master.yaml 에서 계정 코드 집합을 읽는다."""
    try:
        with open(_YAML_PATH, encoding="utf-8") as f:
            data = yaml.safe_load(f)
        return {str(a["code"]) for a in data.get("accounts", [])}
    except Exception:
        # yaml 로드 실패 시 빈 집합 반환 → 모든 코드가 미등록으로 표시되지 않도록 None 처리
        return set()


# 모듈 로드 시 1회만 읽음
ALL_ACCOUNT_CODES: set[str] = _load_account_codes()


class UnknownAccountCodeRule(BaseRule):
    """계정 코드 유효성 검사 — 본사 마스터에 없는 코드 탐지."""

    rule_id = "AC-001"
    name = "계정 코드 유효성 검사"

    def validate(self, tb: TrialBalance, result: ValidationResult) -> None:
        if not ALL_ACCOUNT_CODES:
            return  # yaml 로드 실패 시 규칙 건너뜀

        for row in tb.rows:
            if row.account_code not in ALL_ACCOUNT_CODES:
                result.add(ValidationIssue(
                    rule_id=self.rule_id,
                    severity=Severity.ERROR,
                    message=(
                        f"미등록 계정 코드: [{row.account_code}] {row.account_name} — "
                        "config/accounts_master.yaml 에 없는 코드입니다."
                    ),
                    account_code=row.account_code,
                ))
