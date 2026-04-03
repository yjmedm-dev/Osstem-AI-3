from abc import ABC, abstractmethod
from models.trial_balance import TrialBalance
from models.validation_result import ValidationResult


class BaseRule(ABC):
    """모든 검증 규칙의 공통 기반 클래스.

    새 규칙 추가 방법:
        1. 이 클래스를 상속한다.
        2. rule_id, name, severity 클래스 속성을 정의한다.
        3. validate() 메서드를 구현한다.
    """

    rule_id: str = ""
    name: str = ""

    @abstractmethod
    def validate(self, tb: TrialBalance, result: ValidationResult) -> None:
        """검증을 실행하고, 문제 발견 시 result.add() 로 추가한다."""
        ...
