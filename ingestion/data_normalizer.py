from models.trial_balance import TrialBalance
from utils.currency import convert


def normalize(tb: TrialBalance) -> TrialBalance:
    """TrialBalance 내 모든 금액을 원화로 환산하고 정제한다.

    - debit / credit 이 이미 원화라면 exchange_rate=1 로 제출됨
    - original_amount 는 원본 외화 금액으로 보존됨
    """
    for row in tb.rows:
        if row.original_currency != "KRW":
            krw = convert(row.original_amount, row.original_currency, row.exchange_rate)
            # debit/credit 배분: original_amount 부호로 차변/대변 구분
            if row.original_amount >= 0:
                row.debit = krw
                row.credit = 0.0
            else:
                row.debit = 0.0
                row.credit = abs(krw)
    return tb
