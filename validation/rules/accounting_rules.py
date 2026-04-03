from models.trial_balance import TrialBalance
from models.validation_result import ValidationResult, ValidationIssue, Severity
from validation.rules.base_rule import BaseRule

# accounts_master.yaml 과 동기화된 계정 코드 집합
ALL_ACCOUNT_CODES = {
    "1100", "1200", "1201", "1300", "1301", "1302", "1400",
    "2100", "2200", "2300", "2400", "2401", "2402",
    "3100", "3200",
    "4100", "4200",
    "5100", "5200", "5300",
}


class UnknownAccountCodeRule(BaseRule):
    """계정 코드 유효성 검사 — 본사 마스터에 없는 코드 탐지."""

    rule_id = "AC-001"
    name = "계정 코드 유효성 검사"

    def validate(self, tb: TrialBalance, result: ValidationResult) -> None:
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
