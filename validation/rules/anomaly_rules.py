from models.trial_balance import TrialBalance
from models.validation_result import ValidationResult, ValidationIssue, Severity
from validation.rules.base_rule import BaseRule

# 필수 계정 목록 (required=all)
REQUIRED_ACCOUNT_CODES = {"1100", "1200", "2100", "3100", "3200", "4100", "5100"}

# 자산 계정 (차감 계정 제외)
ASSET_ACCOUNT_CODES = {"1100", "1200", "1300", "1400"}


class RequiredAccountZeroRule(BaseRule):
    """AN-001: 필수 계정 잔액 0 탐지."""

    rule_id = "AN-001"
    name = "필수 계정 잔액 0 탐지"

    def validate(self, tb: TrialBalance, result: ValidationResult) -> None:
        for code in REQUIRED_ACCOUNT_CODES:
            row = tb.get_row(code)
            if row is None or row.balance == 0:
                result.add(ValidationIssue(
                    rule_id=self.rule_id,
                    severity=Severity.WARNING,
                    message=f"필수 계정 [{code}] 잔액이 0이거나 누락되어 있습니다.",
                    account_code=code,
                    actual_value=0.0,
                ))


class NegativeAssetRule(BaseRule):
    """AN-002: 자산 계정 음수 잔액."""

    rule_id = "AN-002"
    name = "자산 계정 음수 잔액"

    def validate(self, tb: TrialBalance, result: ValidationResult) -> None:
        for code in ASSET_ACCOUNT_CODES:
            row = tb.get_row(code)
            if row and row.balance < 0:
                result.add(ValidationIssue(
                    rule_id=self.rule_id,
                    severity=Severity.ERROR,
                    message=(
                        f"자산 계정 [{code}] {row.account_name}에 음수 잔액이 있습니다: "
                        f"{row.balance:,.0f}원"
                    ),
                    account_code=code,
                    actual_value=row.balance,
                    expected_value=0.0,
                ))


class RoundingConcentrationRule(BaseRule):
    """AN-003: 비정상적 반올림 집중 탐지."""

    rule_id = "AN-003"
    name = "비정상적 반올림 집중"
    ROUND_THRESHOLD = 0.8   # 80% 이상이 동일 자릿수 반올림이면 경고

    def validate(self, tb: TrialBalance, result: ValidationResult) -> None:
        balances = [abs(r.balance) for r in tb.rows if r.balance != 0]
        if len(balances) < 5:
            return

        rounded = sum(1 for b in balances if b % 1000 == 0)
        ratio = rounded / len(balances)

        if ratio >= self.ROUND_THRESHOLD:
            result.add(ValidationIssue(
                rule_id=self.rule_id,
                severity=Severity.WARNING,
                message=(
                    f"금액의 {ratio:.0%}가 1,000원 단위 반올림 — "
                    "수기 입력 또는 일괄 조정 가능성이 있습니다."
                ),
                actual_value=ratio,
                expected_value=self.ROUND_THRESHOLD,
            ))
