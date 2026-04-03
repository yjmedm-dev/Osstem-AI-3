from dataclasses import dataclass, field
from typing import Optional
from enum import Enum


class Severity(str, Enum):
    CRITICAL = "CRITICAL"
    ERROR    = "ERROR"
    WARNING  = "WARNING"
    INFO     = "INFO"


@dataclass
class ValidationIssue:
    rule_id: str
    severity: Severity
    message: str
    account_code: Optional[str] = None
    actual_value: Optional[float] = None
    expected_value: Optional[float] = None
    detail: Optional[str] = None


@dataclass
class ValidationResult:
    subsidiary_code: str
    period: str
    issues: list[ValidationIssue] = field(default_factory=list)

    def add(self, issue: ValidationIssue) -> None:
        self.issues.append(issue)

    @property
    def has_critical(self) -> bool:
        return any(i.severity == Severity.CRITICAL for i in self.issues)

    @property
    def has_error(self) -> bool:
        return any(i.severity in (Severity.CRITICAL, Severity.ERROR) for i in self.issues)

    @property
    def is_clean(self) -> bool:
        return len(self.issues) == 0

    def summary(self) -> dict:
        counts = {s: 0 for s in Severity}
        for issue in self.issues:
            counts[issue.severity] += 1
        return {s.value: n for s, n in counts.items() if n > 0}
