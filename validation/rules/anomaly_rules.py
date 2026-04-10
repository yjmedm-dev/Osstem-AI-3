from models.trial_balance import TrialBalance
from models.validation_result import ValidationResult, ValidationIssue, Severity
from validation.rules.base_rule import BaseRule

# 필수 계정 목록 (required=all) — 구 코드 + 신계정 코드 병행 지원
REQUIRED_ACCOUNT_CODES = {
    # 구 계정 코드 체계
    "1100", "1200", "2100", "3100", "3200", "4100", "5100",
    # 신계정 코드 체계 (FP/PL)
    "FP01-01-01-0010",   # 현금및현금성자산
    "FP01-01-01-0070",   # 매출채권
    "FP02-01-01-0010",   # 매입채무
    "FP03-01-01-0010",   # 보통주자본금
    "FP03-05-01-0060",   # 이익잉여금
    "PL01-01-0020",      # 본사제품매출액
    "PL02-01-0020",      # 본사제품매출원가
}

# 자산 계정 (차감 계정 제외) — 구 코드 + 신계정 코드 병행 지원
ASSET_ACCOUNT_CODES = {
    # 구 계정 코드 체계
    "1100", "1200", "1300", "1400",
    # 신계정 코드 체계
    "FP01-01-01-0010",   # 현금및현금성자산
    "FP01-01-01-0070",   # 매출채권
    "FP01-01-02-0010-02", # 본사제품매입분
    "FP01-01-02-0110",   # 저장품
    "FP01-02-02-0110",   # 차량운반구
    "FP01-02-02-0160",   # 집기비품
}


def _detect_code_style(tb: TrialBalance) -> str:
    """시산표 계정 코드 형식 감지: 'new'(FP/PL 신계정) 또는 'legacy'(숫자 구 계정)."""
    for row in tb.rows:
        if row.account_code.startswith(("FP", "PL")):
            return "new"
    return "legacy"


# 구 계정 코드 필수 목록
_REQUIRED_LEGACY = {"1100", "1200", "2100", "3100", "3200", "4100", "5100"}
# 신계정 코드 필수 목록
_REQUIRED_NEW = {
    "FP01-01-01-0010",   # 현금및현금성자산
    "FP01-01-01-0070",   # 매출채권
    "FP02-01-01-0010",   # 매입채무
    "FP03-01-01-0010",   # 보통주자본금
    "FP03-05-01-0060",   # 이익잉여금
    "PL01-01-0020",      # 본사제품매출액
    "PL02-01-0020",      # 본사제품매출원가
}


class RequiredAccountZeroRule(BaseRule):
    """AN-001: 필수 계정 잔액 0 탐지."""

    rule_id = "AN-001"
    name = "필수 계정 잔액 0 탐지"

    def validate(self, tb: TrialBalance, result: ValidationResult) -> None:
        codes = _REQUIRED_NEW if _detect_code_style(tb) == "new" else _REQUIRED_LEGACY
        for code in codes:
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
