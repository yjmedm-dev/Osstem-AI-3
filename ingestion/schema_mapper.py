import pandas as pd

# 법인별 원본 컬럼명 → 표준 컬럼명 매핑 테이블
# 법인 추가 시 이 딕셔너리에 항목만 추가하면 된다.
COLUMN_MAPPINGS: dict[str, dict[str, str]] = {
    "CN01": {
        "科目代码": "account_code",
        "科目名称": "account_name",
        "借方":     "debit",
        "贷方":     "credit",
        "原币金额": "original_amount",
        "币种":     "original_currency",
        "汇率":     "exchange_rate",
    },
    "US01": {
        "Acct Code":    "account_code",
        "Acct Name":    "account_name",
        "Debit":        "debit",
        "Credit":       "credit",
        "Orig Amount":  "original_amount",
        "Currency":     "original_currency",
        "Exch Rate":    "exchange_rate",
    },
    # 그 외 법인은 표준 컬럼명을 그대로 사용한다고 가정
}


def map_columns(df: pd.DataFrame, subsidiary_code: str) -> pd.DataFrame:
    """법인별 컬럼명을 표준 컬럼명으로 변환한다.

    매핑 정의가 없는 법인은 df를 그대로 반환한다.
    """
    mapping = COLUMN_MAPPINGS.get(subsidiary_code.upper())
    if mapping:
        df = df.rename(columns=mapping)
    return df
