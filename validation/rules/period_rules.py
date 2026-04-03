from models.trial_balance import TrialBalance
from models.validation_result import ValidationResult, ValidationIssue, Severity
from validation.rules.base_rule import BaseRule
from config.settings import PERIOD_CHANGE_THRESHOLD


class PeriodChangeRule(BaseRule):
    """PR-001: 전기 대비 30% 초과 변동 항목 탐지."""

    rule_id = "PR-001"
    name = "전기 대비 이상 변동"

    def validate(
        self,
        tb: TrialBalance,
        result: ValidationResult,
        prior_tb: TrialBalance | None = None,
    ) -> None:
        if prior_tb is None:
            return

        for row in tb.rows:
            prior_row = prior_tb.get_row(row.account_code)
            if prior_row is None or prior_row.balance == 0:
                continue

            change_rate = (row.balance - prior_row.balance) / abs(prior_row.balance)

            if abs(change_rate) > PERIOD_CHANGE_THRESHOLD:
                result.add(ValidationIssue(
                    rule_id=self.rule_id,
                    severity=Severity.WARNING,
                    message=(
                        f"[{row.account_code}] {row.account_name}: "
                        f"전기({prior_row.balance:,.0f}) → 당기({row.balance:,.0f}), "
                        f"변동률 {change_rate:+.1%}"
                    ),
                    account_code=row.account_code,
                    actual_value=row.balance,
                    expected_value=prior_row.balance,
                ))
