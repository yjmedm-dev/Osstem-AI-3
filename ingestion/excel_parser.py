from pathlib import Path
import pandas as pd
from models.trial_balance import TrialBalance, TrialBalanceRow
from utils.exceptions import InvalidFileFormatError

# 법인이 제출하는 엑셀의 표준 컬럼명 (schema_mapper.py 에서 매핑 후 이 이름으로 통일)
REQUIRED_COLUMNS = {
    "account_code", "account_name", "debit", "credit",
    "original_amount", "original_currency", "exchange_rate",
}


def parse_excel(
    file_path: Path,
    subsidiary_code: str,
    period: str,
    sheet_name: str = 0,
) -> TrialBalance:
    """엑셀 파일을 읽어 TrialBalance 객체로 반환.

    Args:
        file_path: 원본 엑셀 경로 (data/input/ 하위)
        subsidiary_code: 법인 코드
        period: 기간 문자열 (예: "2025-03")
        sheet_name: 읽을 시트 이름 또는 인덱스

    Returns:
        TrialBalance

    Raises:
        InvalidFileFormatError: 필수 컬럼 누락 또는 파일 읽기 실패 시
    """
    try:
        df = pd.read_excel(file_path, sheet_name=sheet_name, dtype=str)
    except Exception as e:
        raise InvalidFileFormatError(f"파일 읽기 실패 [{file_path.name}]: {e}") from e

    df.columns = [str(c).strip().lower().replace(" ", "_") for c in df.columns]

    missing = REQUIRED_COLUMNS - set(df.columns)
    if missing:
        raise InvalidFileFormatError(
            f"필수 컬럼 누락 [{file_path.name}]: {sorted(missing)}\n"
            "schema_mapper.py 에 이 법인의 컬럼 매핑을 추가하세요."
        )

    rows: list[TrialBalanceRow] = []
    for _, row in df.iterrows():
        try:
            tb_row = TrialBalanceRow(
                subsidiary_code=subsidiary_code,
                period=period,
                account_code=str(row["account_code"]).strip(),
                account_name=str(row["account_name"]).strip(),
                debit=float(row["debit"] or 0),
                credit=float(row["credit"] or 0),
                original_amount=float(row["original_amount"] or 0),
                original_currency=str(row["original_currency"]).strip().upper(),
                exchange_rate=float(row["exchange_rate"] or 1),
            )
            rows.append(tb_row)
        except (ValueError, KeyError) as e:
            raise InvalidFileFormatError(
                f"행 파싱 오류 [{file_path.name}] 계정={row.get('account_code')}: {e}"
            ) from e

    return TrialBalance(
        subsidiary_code=subsidiary_code,
        period=period,
        rows=rows,
        source_file=str(file_path),
    )
