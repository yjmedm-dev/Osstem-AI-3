from models.trial_balance import TrialBalance
from models.validation_result import ValidationResult, ValidationIssue, Severity
from validation.rules.base_rule import BaseRule


class BalanceSheetBalanceRule(BaseRule):
    """AR-001: 차변 합계 = 대변 합계."""

    rule_id = "AR-001"
    name = "대차균형 검증"

    def validate(self, tb: TrialBalance, result: ValidationResult) -> None:
        diff = abs(tb.total_debit - tb.total_credit)
        if diff >= 1.0:
            result.add(ValidationIssue(
                rule_id=self.rule_id,
                severity=Severity.CRITICAL,
                message=f"차변({tb.total_debit:,.0f}) ≠ 대변({tb.total_credit:,.0f}), 차액 {diff:,.0f}원",
                actual_value=diff,
                expected_value=0.0,
            ))


class AccountingEquationRule(BaseRule):
    """AR-002: 자산 = 부채 + 자본."""

    rule_id = "AR-002"
    name = "회계등식 검증"

    ASSET_CODES     = {"1100", "1200", "1201", "1300", "1301", "1302", "1400"}
    LIABILITY_CODES = {"2100", "2200", "2300", "2400", "2401", "2402"}
    EQUITY_CODES    = {"3100", "3200"}

    def validate(self, tb: TrialBalance, result: ValidationResult) -> None:
        assets = sum(tb.get_balance(c) for c in self.ASSET_CODES)
        liabilities = sum(tb.get_balance(c) for c in self.LIABILITY_CODES)
        equity = sum(tb.get_balance(c) for c in self.EQUITY_CODES)
        diff = abs(assets - (liabilities + equity))
        if diff >= 1.0:
            result.add(ValidationIssue(
                rule_id=self.rule_id,
                severity=Severity.CRITICAL,
                message=(
                    f"회계등식 불일치: 자산({assets:,.0f}) ≠ "
                    f"부채({liabilities:,.0f}) + 자본({equity:,.0f}), 차액 {diff:,.0f}원"
                ),
                actual_value=diff,
                expected_value=0.0,
            ))


class RetainedEarningsRule(BaseRule):
    """AR-003: 이익잉여금 검증 (전기 이익잉여금 + 당기순이익 = 기말 이익잉여금)."""

    rule_id = "AR-003"
    name = "이익잉여금 검증"

    RETAINED_EARNINGS_CODE = "3200"
    REVENUE_CODE           = "4100"
    COGS_CODE              = "5100"
    SGA_CODE               = "5200"

    def validate(self, tb: TrialBalance, result: ValidationResult) -> None:
        # 간이 검증: 당기순이익 = 수익 - 비용 이 이익잉여금 증감과 일치하는지
        # 전기 데이터가 없으므로 현재는 플래그만 발행
        revenue  = tb.get_balance(self.REVENUE_CODE)
        cogs     = tb.get_balance(self.COGS_CODE)
        sga      = tb.get_balance(self.SGA_CODE)
        net_income = revenue - cogs - sga

        if abs(net_income) > 0 and tb.get_balance(self.RETAINED_EARNINGS_CODE) == 0:
            result.add(ValidationIssue(
                rule_id=self.rule_id,
                severity=Severity.ERROR,
                message="이익잉여금 계정 잔액이 0이지만 당기순이익이 존재합니다. 확인이 필요합니다.",
                account_code=self.RETAINED_EARNINGS_CODE,
                actual_value=0.0,
                expected_value=net_income,
            ))
