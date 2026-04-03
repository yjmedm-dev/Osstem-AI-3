from dataclasses import dataclass
from typing import Optional


@dataclass
class Account:
    code: str
    name: str
    account_type: str        # asset / liability / equity / revenue / expense
    required: str            # all / manufacturing / trading / retirement_mandatory
    is_provision: bool = False
    provision_type: Optional[str] = None   # PV 규칙의 provision_type과 매핑
    ifrs_ref: Optional[str] = None         # 예: "IFRS16", "IFRS15"

    def is_debit_normal(self) -> bool:
        """차변 잔액이 정상인 계정 유형 여부."""
        return self.account_type in ("asset", "expense")

    def is_contra(self) -> bool:
        """충당금처럼 차감 표시되는 계정 여부."""
        return self.is_provision and self.account_type == "asset"
