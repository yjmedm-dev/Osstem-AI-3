from utils.exceptions import CurrencyConversionError

# 기준 환율 테이블 (실제 운용 시 외부 API 또는 data/reference/exchange_rates.yaml 로 대체)
_FALLBACK_RATES: dict[str, float] = {
    "KRW": 1.0,
    "USD": 1340.0,
    "CNY": 185.0,
    "EUR": 1450.0,
    "AUD": 870.0,
    "BRL": 265.0,
    "INR": 16.1,
    "RUB": 14.5,
    "VND": 0.054,
    "IDR": 0.086,
    "JPY": 8.9,
    "UZS": 0.106,   # 우즈베키스탄 숨 (1 USD≈12,600 UZS, 1 USD≈1,340 KRW 기준)
    "THB": 38.5,
}


def convert(amount: float, from_currency: str, rate: float | None = None) -> float:
    """외화 금액을 원화(KRW)로 변환한다.

    Args:
        amount: 외화 금액
        from_currency: 출발 통화 코드 (예: "USD")
        rate: 적용할 환율 (None이면 내부 기본값 사용)

    Returns:
        원화 금액 (float)

    Raises:
        CurrencyConversionError: 지원하지 않는 통화이고 rate도 없는 경우
    """
    currency = from_currency.upper()
    if currency == "KRW":
        return amount

    if rate is None:
        rate = _FALLBACK_RATES.get(currency)
        if rate is None:
            raise CurrencyConversionError(
                f"환율 정보 없음: {currency}. rate 인자를 직접 전달하거나 "
                "data/reference/exchange_rates.yaml 에 추가하세요."
            )

    return amount * rate
