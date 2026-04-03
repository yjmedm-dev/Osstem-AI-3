from dataclasses import dataclass, field
from typing import Optional
from datetime import date


@dataclass
class TrialBalanceRow:
    subsidiary_code: str
    period: str              # 예: "2025-03"
    account_code: str
    account_name: str
    debit: float             # 차변 (원화 KRW)
    credit: float            # 대변 (원화 KRW)
    original_amount: float   # 외화 원본 금액
    original_currency: str   # 외화 통화 코드
    exchange_rate: float     # 적용 환율

    @property
    def balance(self) -> float:
        """순잔액 (차변 − 대변)."""
        return self.debit - self.credit


@dataclass
class TrialBalance:
    subsidiary_code: str
    period: str
    rows: list[TrialBalanceRow] = field(default_factory=list)
    source_file: Optional[str] = None

    @property
    def total_debit(self) -> float:
        return sum(r.debit for r in self.rows)

    @property
    def total_credit(self) -> float:
        return sum(r.credit for r in self.rows)

    @property
    def is_balanced(self) -> bool:
        return abs(self.total_debit - self.total_credit) < 1.0  # 1원 미만 오차 허용

    def get_row(self, account_code: str) -> Optional[TrialBalanceRow]:
        for row in self.rows:
            if row.account_code == account_code:
                return row
        return None

    def get_balance(self, account_code: str) -> float:
        row = self.get_row(account_code)
        return row.balance if row else 0.0
