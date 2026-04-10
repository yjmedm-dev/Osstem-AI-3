from pathlib import Path
import pandas as pd
from models.trial_balance import TrialBalance, TrialBalanceRow
from utils.currency import convert
from utils.exceptions import InvalidFileFormatError
from ingestion.schema_mapper import map_columns

# 법인이 제출하는 엑셀의 표준 컬럼명 (schema_mapper.py 에서 매핑 후 이 이름으로 통일)
REQUIRED_COLUMNS = {
    "account_code", "account_name", "debit", "credit",
    "original_amount", "original_currency", "exchange_rate",
}


def parse_excel(
    file_path: Path,
    subsidiary_code: str,
    period: str,
    sheet_name: str | int = 0,
    exchange_rate: float | None = None,
) -> TrialBalance:
    """엑셀 파일을 읽어 TrialBalance 객체로 반환.

    UZ01(우즈베키스탄)은 별도 전용 파서로 처리한다.

    Args:
        file_path: 원본 엑셀 경로 (data/input/ 하위)
        subsidiary_code: 법인 코드
        period: 기간 문자열 (예: "2025-03")
        sheet_name: 읽을 시트 이름 또는 인덱스
        exchange_rate: 외화→KRW 환율 (None이면 폴백 환율 사용)

    Returns:
        TrialBalance

    Raises:
        InvalidFileFormatError: 필수 컬럼 누락 또는 파일 읽기 실패 시
    """
    if subsidiary_code.upper() == "UZ01":
        return parse_uz01_excel(file_path, period, exchange_rate=exchange_rate)

    try:
        df = pd.read_excel(file_path, sheet_name=sheet_name, dtype=str)
    except Exception as e:
        raise InvalidFileFormatError(f"파일 읽기 실패 [{file_path.name}]: {e}") from e

    df = map_columns(df, subsidiary_code)
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


def parse_uz01_excel(
    file_path: Path,
    period: str,
    exchange_rate: float | None = None,
) -> TrialBalance:
    """우즈베키스탄(UZ01) 마감자료 엑셀 → TrialBalance.

    '전체' 시트를 파싱하여 한국 신계정 코드 기준으로 집계한다.
    차변·대변은 기말잔액(신계정별 합산)을 사용한다.

    Args:
        file_path: 우즈벡 마감자료 엑셀 경로
        period: 기간 문자열 (예: "2026-02")
        exchange_rate: 1 UZS → KRW 환율 (None이면 폴백 0.106 사용)

    Returns:
        TrialBalance (debit/credit 단위: KRW)
    """
    try:
        raw = pd.read_excel(file_path, sheet_name="전체", header=4)
    except Exception as e:
        raise InvalidFileFormatError(
            f"파일 읽기 실패 [{file_path.name}] 시트='전체': {e}"
        ) from e

    # 컬럼 순서(헤더 row=4 기준):
    #   0:구분  1:소계정  2:우즈벡계정
    #   3:기초Д  4:기초К  5:오버턴Д  6:오버턴К  7:기말Д  8:기말К
    #   9:비고  10:신계정코드  11:신계정  12:차변(기말집계)  13:대변(기말집계)
    cols = list(raw.columns)
    if len(cols) < 14:
        raise InvalidFileFormatError(
            f"[전체] 시트 컬럼 수 부족: {len(cols)}개 (최소 14개 필요)"
        )

    raw = raw.rename(columns={
        cols[2]:  "uz_account",
        cols[10]: "std_account_code",
        cols[11]: "account_name",
        cols[12]: "debit_uzs",
        cols[13]: "credit_uzs",
    })

    # 신계정 코드가 있는 리프 노드 행만 추출
    data = raw[
        raw["std_account_code"].notna()
        & raw["std_account_code"].astype(str).str.strip().ne("")
    ].copy()

    if data.empty:
        raise InvalidFileFormatError(
            f"[전체] 시트에서 신계정 코드가 매핑된 행을 찾을 수 없습니다: {file_path.name}"
        )

    data["debit_uzs"]  = pd.to_numeric(data["debit_uzs"],  errors="coerce").fillna(0.0)
    data["credit_uzs"] = pd.to_numeric(data["credit_uzs"], errors="coerce").fillna(0.0)

    # 신계정 코드별로 차변·대변 합산 (여러 1C 계정 → 하나의 신계정)
    grouped = (
        data.groupby(["std_account_code", "account_name"], as_index=False)
        .agg(debit_uzs=("debit_uzs", "sum"), credit_uzs=("credit_uzs", "sum"))
    )

    rows: list[TrialBalanceRow] = []
    for _, row in grouped.iterrows():
        debit_krw  = convert(row["debit_uzs"],  "UZS", exchange_rate)
        credit_krw = convert(row["credit_uzs"], "UZS", exchange_rate)
        uzs_net = row["debit_uzs"] - row["credit_uzs"]
        uzs_rate = exchange_rate if exchange_rate is not None else 0.106

        rows.append(TrialBalanceRow(
            subsidiary_code="UZ01",
            period=period,
            account_code=str(row["std_account_code"]).strip(),
            account_name=str(row["account_name"]).strip(),
            debit=debit_krw,
            credit=credit_krw,
            original_amount=uzs_net,
            original_currency="UZS",
            exchange_rate=uzs_rate,
        ))

    return TrialBalance(
        subsidiary_code="UZ01",
        period=period,
        rows=rows,
        source_file=str(file_path),
    )
