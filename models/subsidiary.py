from dataclasses import dataclass, field
from typing import Optional


@dataclass
class Subsidiary:
    code: str                          # 예: "CN01"
    name: str                          # 예: "오스템 중국 법인"
    country: str                       # 예: "CN"
    currency: str                      # 기능통화 예: "CNY"
    entity_type: str                   # "manufacturing" | "trading"
    retirement_mandatory: bool = False # 퇴직급여충당금 의무 여부
    active: bool = True

    def __post_init__(self):
        self.code = self.code.upper()
        self.currency = self.currency.upper()
