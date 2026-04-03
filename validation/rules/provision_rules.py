"""충당금 전용 검증 규칙 (PV-001 ~ PV-009).

충당금 8종 각각에 대해 BaseRule을 상속하여 구현한다.
"""

from models.trial_balance import TrialBalance
from models.validation_result import ValidationResult, ValidationIssue, Severity
from validation.rules.base_rule import BaseRule
from config.settings import RETIREMENT_MANDATORY_SUBSIDIARIES


def _pct_diff(actual: float, expected: float) -> float:
    """expected 대비 actual의 오차율 (절대값 기준)."""
    if expected == 0:
        return 0.0
    return abs(actual - expected) / abs(expected)


class InventoryValuationProvisionRule(BaseRule):
    """PV-001: 재고자산평가충당금 검증.

    설정액 = 저속·불용 재고 잔액 × 요율
    - 재고 계정(1300) 대비 충당금(1301) 비율이 설정 요율 범위 내인지 확인
    """

    rule_id = "PV-001"
    name = "재고자산평가충당금 검증"
    INVENTORY_CODE   = "1300"
    PROVISION_CODE   = "1301"
    MIN_RATE = 0.02   # 최소 설정 요율 2%
    MAX_RATE = 0.30   # 최대 설정 요율 30%
    TOLERANCE = 0.01

    def validate(self, tb: TrialBalance, result: ValidationResult) -> None:
        inventory  = tb.get_balance(self.INVENTORY_CODE)
        provision  = abs(tb.get_balance(self.PROVISION_CODE))

        if inventory == 0:
            return

        rate = provision / inventory
        if rate < self.MIN_RATE:
            result.add(ValidationIssue(
                rule_id=self.rule_id,
                severity=Severity.ERROR,
                message=(
                    f"재고자산평가충당금 설정 요율({rate:.1%})이 최소 기준({self.MIN_RATE:.0%}) 미달. "
                    f"재고({inventory:,.0f}), 충당금({provision:,.0f})"
                ),
                account_code=self.PROVISION_CODE,
                actual_value=rate,
                expected_value=self.MIN_RATE,
            ))
        elif rate > self.MAX_RATE:
            result.add(ValidationIssue(
                rule_id=self.rule_id,
                severity=Severity.WARNING,
                message=(
                    f"재고자산평가충당금 설정 요율({rate:.1%})이 최대 기준({self.MAX_RATE:.0%}) 초과 — 검토 필요"
                ),
                account_code=self.PROVISION_CODE,
                actual_value=rate,
                expected_value=self.MAX_RATE,
            ))


class DoubtfulAccountsProvisionRule(BaseRule):
    """PV-002: 대손충당금 검증.

    설정액 = 매출채권 잔액 × 설정 요율
    - 채권(1200) 대비 충당금(1201) 비율 검증
    """

    rule_id = "PV-002"
    name = "대손충당금 검증"
    RECEIVABLE_CODE = "1200"
    PROVISION_CODE  = "1201"
    MIN_RATE = 0.005   # 최소 0.5%
    MAX_RATE = 0.50    # 최대 50%
    TOLERANCE = 0.01

    def validate(self, tb: TrialBalance, result: ValidationResult) -> None:
        receivable = tb.get_balance(self.RECEIVABLE_CODE)
        provision  = abs(tb.get_balance(self.PROVISION_CODE))

        if receivable == 0:
            return

        if provision == 0:
            result.add(ValidationIssue(
                rule_id=self.rule_id,
                severity=Severity.ERROR,
                message=(
                    f"매출채권({receivable:,.0f}원)이 있으나 대손충당금이 미설정되어 있습니다."
                ),
                account_code=self.PROVISION_CODE,
                actual_value=0.0,
                expected_value=receivable * self.MIN_RATE,
            ))
            return

        rate = provision / receivable
        if rate < self.MIN_RATE:
            result.add(ValidationIssue(
                rule_id=self.rule_id,
                severity=Severity.WARNING,
                message=(
                    f"대손충당금 설정 요율({rate:.2%})이 최소 기준({self.MIN_RATE:.1%}) 미달."
                ),
                account_code=self.PROVISION_CODE,
                actual_value=rate,
                expected_value=self.MIN_RATE,
            ))


class RetirementBenefitProvisionRule(BaseRule):
    """PV-003: 퇴직급여충당금 검증.

    의무 있는 법인에만 적용.
    충당금(2300) 잔액이 0이면 ERROR.
    """

    rule_id = "PV-003"
    name = "퇴직급여충당금 검증"
    PROVISION_CODE = "2300"

    def validate(self, tb: TrialBalance, result: ValidationResult) -> None:
        if tb.subsidiary_code not in RETIREMENT_MANDATORY_SUBSIDIARIES:
            return

        provision = tb.get_balance(self.PROVISION_CODE)
        if provision == 0:
            result.add(ValidationIssue(
                rule_id=self.rule_id,
                severity=Severity.ERROR,
                message=(
                    f"[{tb.subsidiary_code}] 퇴직급여충당금 의무 법인이나 충당금이 미설정되어 있습니다."
                ),
                account_code=self.PROVISION_CODE,
                actual_value=0.0,
            ))


class LeaseLiabilityProvisionRule(BaseRule):
    """PV-004: 리스회계충당금 검증 (IFRS 16).

    ROU자산(1400)이 있는 경우 리스부채(2200)도 존재해야 한다.
    """

    rule_id = "PV-004"
    name = "리스회계충당금 검증"
    ROU_CODE   = "1400"
    LEASE_CODE = "2200"

    def validate(self, tb: TrialBalance, result: ValidationResult) -> None:
        rou   = tb.get_balance(self.ROU_CODE)
        lease = abs(tb.get_balance(self.LEASE_CODE))

        if rou > 0 and lease == 0:
            result.add(ValidationIssue(
                rule_id=self.rule_id,
                severity=Severity.ERROR,
                message=(
                    f"ROU자산({rou:,.0f}원)이 있으나 리스부채가 미설정되어 있습니다. "
                    "IFRS 16 적용 여부를 확인하세요."
                ),
                account_code=self.LEASE_CODE,
                actual_value=0.0,
                expected_value=rou,
            ))


class SalesReturnProvisionRule(BaseRule):
    """PV-005: 반품충당금 검증.

    매출(4100)이 있는 경우 반품충당금(2400) 존재 여부 확인.
    """

    rule_id = "PV-005"
    name = "반품충당금 검증"
    REVENUE_CODE   = "4100"
    PROVISION_CODE = "2400"
    MIN_RATE = 0.001   # 최소 0.1%
    TOLERANCE = 0.05

    def validate(self, tb: TrialBalance, result: ValidationResult) -> None:
        revenue   = tb.get_balance(self.REVENUE_CODE)
        provision = abs(tb.get_balance(self.PROVISION_CODE))

        if revenue == 0:
            return

        if provision == 0:
            result.add(ValidationIssue(
                rule_id=self.rule_id,
                severity=Severity.WARNING,
                message=(
                    f"매출({revenue:,.0f}원)이 있으나 반품충당금이 미설정되어 있습니다."
                ),
                account_code=self.PROVISION_CODE,
                actual_value=0.0,
                expected_value=revenue * self.MIN_RATE,
            ))


class DiscontinuedItemProvisionRule(BaseRule):
    """PV-006: 단품충당금 검증.

    재고(1300)가 있는 경우 단품충당금(1302) 설정 여부 확인.
    """

    rule_id = "PV-006"
    name = "단품충당금 검증"
    INVENTORY_CODE = "1300"
    PROVISION_CODE = "1302"

    def validate(self, tb: TrialBalance, result: ValidationResult) -> None:
        inventory = tb.get_balance(self.INVENTORY_CODE)
        provision = abs(tb.get_balance(self.PROVISION_CODE))

        if inventory > 0 and provision == 0:
            result.add(ValidationIssue(
                rule_id=self.rule_id,
                severity=Severity.WARNING,
                message=(
                    f"재고({inventory:,.0f}원)가 있으나 단품충당금이 미설정되어 있습니다. "
                    "단종·사양화 SKU 해당 여부를 확인하세요."
                ),
                account_code=self.PROVISION_CODE,
                actual_value=0.0,
            ))


class FocProvisionRule(BaseRule):
    """PV-007: FOC충당금 검증.

    매출(4100)이 있는 경우 FOC충당금(2401) 존재 여부 확인.
    """

    rule_id = "PV-007"
    name = "FOC충당금 검증"
    REVENUE_CODE   = "4100"
    PROVISION_CODE = "2401"

    def validate(self, tb: TrialBalance, result: ValidationResult) -> None:
        revenue   = tb.get_balance(self.REVENUE_CODE)
        provision = abs(tb.get_balance(self.PROVISION_CODE))

        if revenue > 0 and provision == 0:
            result.add(ValidationIssue(
                rule_id=self.rule_id,
                severity=Severity.WARNING,
                message=(
                    f"매출({revenue:,.0f}원)이 있으나 FOC충당금이 미설정되어 있습니다. "
                    "무상공급(FOC) 정책 적용 여부를 확인하세요."
                ),
                account_code=self.PROVISION_CODE,
                actual_value=0.0,
            ))


class RevenueRecognitionProvisionRule(BaseRule):
    """PV-008: 수익인식충당금 검증 (IFRS 15).

    매출(4100)이 있는 경우 수익인식충당금(2402) 존재 여부 확인.
    """

    rule_id = "PV-008"
    name = "수익인식충당금 검증"
    REVENUE_CODE   = "4100"
    PROVISION_CODE = "2402"

    def validate(self, tb: TrialBalance, result: ValidationResult) -> None:
        revenue   = tb.get_balance(self.REVENUE_CODE)
        provision = abs(tb.get_balance(self.PROVISION_CODE))

        if revenue > 0 and provision == 0:
            result.add(ValidationIssue(
                rule_id=self.rule_id,
                severity=Severity.WARNING,
                message=(
                    f"매출({revenue:,.0f}원)이 있으나 수익인식충당금이 미설정되어 있습니다. "
                    "IFRS 15 조건부 수익 이연 여부를 확인하세요."
                ),
                account_code=self.PROVISION_CODE,
                actual_value=0.0,
            ))


class ProvisionReversalRule(BaseRule):
    """PV-009: 충당금 환입 패턴 탐지.

    전기 대비 충당금 잔액이 감소한 경우 경고 발행.
    """

    rule_id = "PV-009"
    name = "충당금 환입 사유 체크"

    PROVISION_CODES = {"1201", "1301", "1302", "2200", "2300", "2400", "2401", "2402"}

    def validate(
        self,
        tb: TrialBalance,
        result: ValidationResult,
        prior_tb: TrialBalance | None = None,
    ) -> None:
        if prior_tb is None:
            return

        for code in self.PROVISION_CODES:
            current  = abs(tb.get_balance(code))
            previous = abs(prior_tb.get_balance(code))

            if previous > 0 and current < previous:
                reduction = previous - current
                result.add(ValidationIssue(
                    rule_id=self.rule_id,
                    severity=Severity.WARNING,
                    message=(
                        f"충당금 [{code}] 환입 감지: 전기({previous:,.0f}) → 당기({current:,.0f}), "
                        f"감소액 {reduction:,.0f}원 — 환입 사유를 확인하세요."
                    ),
                    account_code=code,
                    actual_value=current,
                    expected_value=previous,
                ))
