from models.trial_balance import TrialBalance
from models.validation_result import ValidationResult
from validation.rules.arithmetic_rules import (
    BalanceSheetBalanceRule,
    AccountingEquationRule,
    RetainedEarningsRule,
)
from validation.rules.accounting_rules import UnknownAccountCodeRule
from validation.rules.period_rules import PeriodChangeRule
from validation.rules.anomaly_rules import (
    RequiredAccountZeroRule,
    NegativeAssetRule,
    RoundingConcentrationRule,
)
from validation.rules.provision_rules import (
    InventoryValuationProvisionRule,
    DoubtfulAccountsProvisionRule,
    RetirementBenefitProvisionRule,
    LeaseLiabilityProvisionRule,
    SalesReturnProvisionRule,
    DiscontinuedItemProvisionRule,
    FocProvisionRule,
    RevenueRecognitionProvisionRule,
    ProvisionReversalRule,
)


class ValidationEngine:
    """모든 검증 규칙을 순서대로 실행하는 엔진."""

    def __init__(self):
        # CRITICAL → ERROR → WARNING 순으로 등록
        self._rules = [
            BalanceSheetBalanceRule(),
            AccountingEquationRule(),
            RetainedEarningsRule(),
            UnknownAccountCodeRule(),
            RequiredAccountZeroRule(),
            NegativeAssetRule(),
            RoundingConcentrationRule(),
            PeriodChangeRule(),
            InventoryValuationProvisionRule(),
            DoubtfulAccountsProvisionRule(),
            RetirementBenefitProvisionRule(),
            LeaseLiabilityProvisionRule(),
            SalesReturnProvisionRule(),
            DiscontinuedItemProvisionRule(),
            FocProvisionRule(),
            RevenueRecognitionProvisionRule(),
            ProvisionReversalRule(),
        ]

    def run(
        self,
        tb: TrialBalance,
        prior_tb: TrialBalance | None = None,
    ) -> ValidationResult:
        """검증 실행 후 ValidationResult 반환.

        Args:
            tb: 검증 대상 시산표
            prior_tb: 전기 시산표 (없으면 기간 비교·환입 규칙 건너뜀)
        """
        result = ValidationResult(
            subsidiary_code=tb.subsidiary_code,
            period=tb.period,
        )

        for rule in self._rules:
            try:
                # 전기 데이터가 필요한 규칙은 prior_tb 인자를 추가로 전달
                if isinstance(rule, (PeriodChangeRule, ProvisionReversalRule)):
                    rule.validate(tb, result, prior_tb=prior_tb)
                else:
                    rule.validate(tb, result)
            except Exception as e:
                # 규칙 자체 오류는 검증을 중단하지 않고 로그만 남긴다
                from models.validation_result import ValidationIssue, Severity
                result.add(ValidationIssue(
                    rule_id=rule.rule_id,
                    severity=Severity.ERROR,
                    message=f"규칙 실행 오류 [{rule.rule_id}]: {e}",
                ))

        return result
