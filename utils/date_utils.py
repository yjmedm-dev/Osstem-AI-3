from datetime import date
from dateutil.relativedelta import relativedelta


def period_to_date(period: str) -> date:
    """'2025-03' 형식의 기간 문자열을 해당 월 말일 date로 변환."""
    year, month = map(int, period.split("-"))
    first_day = date(year, month, 1)
    last_day = first_day + relativedelta(months=1) - relativedelta(days=1)
    return last_day


def prior_period(period: str) -> str:
    """전월 기간 문자열 반환. 예: '2025-03' → '2025-02'."""
    year, month = map(int, period.split("-"))
    d = date(year, month, 1) - relativedelta(months=1)
    return d.strftime("%Y-%m")


def prior_year_period(period: str) -> str:
    """전년 동월 기간 문자열 반환. 예: '2025-03' → '2024-03'."""
    year, month = map(int, period.split("-"))
    return f"{year - 1:04d}-{month:02d}"


def periods_in_year(year: int) -> list[str]:
    """해당 연도의 월별 기간 목록 반환."""
    return [f"{year:04d}-{m:02d}" for m in range(1, 13)]


def remaining_periods(period: str) -> list[str]:
    """당월 포함 연말까지 남은 기간 목록 반환."""
    year, month = map(int, period.split("-"))
    return [f"{year:04d}-{m:02d}" for m in range(month, 13)]
