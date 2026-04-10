from models.trial_balance import TrialBalance
from models.validation_result import ValidationResult, ValidationIssue, Severity
from validation.rules.base_rule import BaseRule

# 당기순손익 계정 코드 (신계정 / 구계정)
_NET_INCOME_CODES = {"FP03-05-01-0070", "3200"}


def _net_income_balance(tb: TrialBalance) -> float:
    """시산표에서 당기순손익 계정의 순잔액을 반환한다.

    잔액시산표에서 당기순손실이면 차변 잔액(양수), 당기순이익이면 대변 잔액(음수).
    """
    for code in _NET_INCOME_CODES:
        row = tb.get_row(code)
        if row is not None:
            return row.balance  # debit - credit
    return 0.0


class BalanceSheetBalanceRule(BaseRule):
    """AR-001: 차변 합계 = 대변 합계.

    잔액시산표(Trial Balance)의 정상 특성:
        차대 불균형 = 당기순손익 계정 잔액  →  정상 (INFO)
        차대 불균형 ≠ 당기순손익 계정 잔액  →  실제 오류 (CRITICAL)
    """

    rule_id = "AR-001"
    name = "대차균형 검증"

    def validate(self, tb: TrialBalance, result: ValidationResult) -> None:
        diff = tb.total_debit - tb.total_credit          # 부호 포함
        abs_diff = abs(diff)

        if abs_diff < 1.0:
            return  # 완전 균형 → 이슈 없음

        net_income = _net_income_balance(tb)             # 당기순손익 계정 잔액
        abs_net = abs(net_income)

        if abs(abs_diff - abs_net) < 1.0:
            # 불균형 = 당기순손익 → 잔액시산표 정상 특성
            result.add(ValidationIssue(
                rule_id=self.rule_id,
                severity=Severity.INFO,
                message=(
                    f"차대 불균형({abs_diff:,.0f})이 당기순손익 계정 잔액({abs_net:,.0f})과 "
                    f"일치 — 잔액시산표 정상 특성. "
                    f"{'당기순손실' if diff > 0 else '당기순이익'}: {abs_net:,.0f}"
                ),
                actual_value=abs_diff,
                expected_value=abs_net,
            ))
        else:
            # 설명되지 않는 차대 불균형 → 실제 오류
            result.add(ValidationIssue(
                rule_id=self.rule_id,
                severity=Severity.CRITICAL,
                message=(
                    f"차변({tb.total_debit:,.0f}) ≠ 대변({tb.total_credit:,.0f}), "
                    f"차액 {abs_diff:,.0f} "
                    f"(당기순손익 {abs_net:,.0f}으로 설명 불가 — 분개 오류 의심)"
                ),
                actual_value=abs_diff,
                expected_value=abs_net,
            ))


class AccountingEquationRule(BaseRule):
    """AR-002: 자산 = 부채 + 자본.

    신계정 코드(FP 접두사)와 구계정 코드(숫자) 모두 지원.
    """

    rule_id = "AR-002"
    name = "회계등식 검증"

    # 구계정 코드
    _LEGACY_ASSET     = {"1100", "1200", "1201", "1300", "1301", "1302", "1400"}
    _LEGACY_LIABILITY = {"2100", "2200", "2300", "2400", "2401", "2402"}
    _LEGACY_EQUITY    = {"3100", "3200"}

    def validate(self, tb: TrialBalance, result: ValidationResult) -> None:
        # 신계정 코드 여부 자동 감지
        uses_new = any(r.account_code.startswith(("FP", "PL")) for r in tb.rows)

        if uses_new:
            # FP01-xx: 자산 / FP02-xx: 부채 / FP03-xx: 자본
            assets      = sum(r.balance for r in tb.rows if r.account_code.startswith("FP01"))
            liabilities = sum(r.balance for r in tb.rows if r.account_code.startswith("FP02"))
            equity      = sum(r.balance for r in tb.rows if r.account_code.startswith("FP03"))
        else:
            assets      = sum(tb.get_balance(c) for c in self._LEGACY_ASSET)
            liabilities = sum(tb.get_balance(c) for c in self._LEGACY_LIABILITY)
            equity      = sum(tb.get_balance(c) for c in self._LEGACY_EQUITY)

        if uses_new:
            # 신계정: balance = debit - credit
            #   자산(FP01): 차변 잔액 → 양수
            #   부채(FP02): 대변 잔액 → 음수
            #   자본(FP03): 대변 잔액 → 음수
            # 균형 조건: 자산 + 부채 + 자본 = 0
            diff = abs(assets + liabilities + equity)
            liab_display  = abs(liabilities)
            equity_display = abs(equity)
        else:
            diff = abs(assets - (liabilities + equity))
            liab_display  = liabilities
            equity_display = equity

        if diff >= 1.0:
            result.add(ValidationIssue(
                rule_id=self.rule_id,
                severity=Severity.CRITICAL,
                message=(
                    f"회계등식 불일치: 자산({assets:,.0f}) ≠ "
                    f"부채({liab_display:,.0f}) + 자본({equity_display:,.0f}), 차액 {diff:,.0f}"
                ),
                actual_value=diff,
                expected_value=0.0,
            ))


class RetainedEarningsRule(BaseRule):
    """AR-003: 이익잉여금 검증."""

    rule_id = "AR-003"
    name = "이익잉여금 검증"

    # 구계정
    _LEGACY_RETAINED = "3200"
    _LEGACY_REVENUE  = "4100"
    _LEGACY_COGS     = "5100"
    _LEGACY_SGA      = "5200"

    # 신계정
    _NEW_RETAINED_CURR = "FP03-05-01-0070"   # 미처분이익잉여금_당기순손익
    _NEW_REVENUE       = "PL01-01-0020"      # 대표 매출 계정
    _NEW_COGS          = "PL02-01-0020"      # 대표 원가 계정

    def validate(self, tb: TrialBalance, result: ValidationResult) -> None:
        uses_new = any(r.account_code.startswith(("FP", "PL")) for r in tb.rows)

        if uses_new:
            net_income_row = tb.get_row(self._NEW_RETAINED_CURR)
            retained = net_income_row.balance if net_income_row else 0.0
            revenue  = -tb.get_balance(self._NEW_REVENUE)   # 수익은 대변 → 음수 잔액
            cogs     =  tb.get_balance(self._NEW_COGS)
        else:
            retained = tb.get_balance(self._LEGACY_RETAINED)
            revenue  = tb.get_balance(self._LEGACY_REVENUE)
            cogs     = tb.get_balance(self._LEGACY_COGS) + tb.get_balance(self._LEGACY_SGA)

        # 당기순손익이 있는데 이익잉여금이 0이면 경고
        net_income = revenue - cogs
        if abs(net_income) > 0 and retained == 0:
            result.add(ValidationIssue(
                rule_id=self.rule_id,
                severity=Severity.ERROR,
                message="이익잉여금(당기) 계정 잔액이 0이지만 당기순손익이 존재합니다.",
                account_code=self._NEW_RETAINED_CURR if uses_new else self._LEGACY_RETAINED,
                actual_value=0.0,
                expected_value=net_income,
            ))
