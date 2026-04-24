"""현지회계 데이터 + 마스터 매핑 → Confinas 업로드용 엑셀 생성

Confinas 양식이 확정되면 _write_confinas_sheet() 내부의 컬럼 순서와
헤더명만 수정하면 된다. 집계 로직은 그대로 재사용된다.
"""
from pathlib import Path

import xlsxwriter
from sqlalchemy import select

from db.connection import get_session
from db.models import AccountMaster, FinancialLocal


def export(
    subsidiary_code: str,
    period: str,
    out_filepath: str | Path,
    template_columns: list[str] | None = None,
) -> int:
    """Confinas 업로드용 엑셀을 생성한다.

    Args:
        subsidiary_code: 법인 코드 (예: 'UZ01')
        period: 기간 (예: '2025-03')
        out_filepath: 출력 파일 경로
        template_columns: Confinas 양식의 컬럼 순서 리스트.
            None이면 기본 양식(아래 _DEFAULT_COLUMNS)을 사용.

    Returns:
        출력된 행 수 (헤더 제외)
    """
    subsidiary_code = subsidiary_code.upper()
    out_filepath    = Path(out_filepath)

    # ── 데이터 조회 ──────────────────────────────────────────────────────────
    with get_session() as session:
        master_rows = session.execute(
            select(AccountMaster).where(
                AccountMaster.subsidiary_code == subsidiary_code,
                AccountMaster.confinas_code.isnot(None),
            )
        ).scalars().all()

        local_rows = session.execute(
            select(FinancialLocal).where(
                FinancialLocal.subsidiary_code == subsidiary_code,
                FinancialLocal.period == period,
            )
        ).scalars().all()

    # ── local_code → 잔액 매핑 ───────────────────────────────────────────────
    local_data: dict[str, dict] = {}
    for r in local_rows:
        local_data[r.account_code] = {
            "account_name": r.account_name or "",
            "debit":        float(r.debit or 0),
            "credit":       float(r.credit or 0),
            "balance":      float(r.balance or 0),
            "currency":     r.currency or "",
            "amount_krw":   float(r.amount_krw or 0),
        }

    # ── confinas_code 기준으로 합산 ───────────────────────────────────────────
    confinas_agg: dict[str, dict] = {}
    local_to_confinas: dict[str, str] = {}

    for m in master_rows:
        cc = m.confinas_code
        lc = m.local_code or ""
        local_to_confinas[lc] = cc

        if cc not in confinas_agg:
            confinas_agg[cc] = {
                "confinas_code": cc,
                "confinas_name": m.confinas_name or "",
                "debit":         0.0,
                "credit":        0.0,
                "balance":       0.0,
                "amount_krw":    0.0,
                "currency":      "",
            }

        d = local_data.get(lc, {})
        confinas_agg[cc]["debit"]      += d.get("debit",      0.0)
        confinas_agg[cc]["credit"]     += d.get("credit",     0.0)
        confinas_agg[cc]["balance"]    += d.get("balance",    0.0)
        confinas_agg[cc]["amount_krw"] += d.get("amount_krw", 0.0)
        if not confinas_agg[cc]["currency"]:
            confinas_agg[cc]["currency"] = d.get("currency", "")

    rows = sorted(confinas_agg.values(), key=lambda x: x["confinas_code"])

    # ── 엑셀 출력 ────────────────────────────────────────────────────────────
    _write_confinas_sheet(rows, subsidiary_code, period, out_filepath, template_columns)
    return len(rows)


# ── 기본 Confinas 컬럼 양식 ──────────────────────────────────────────────────
# Confinas 실제 양식이 확정되면 이 목록을 맞게 수정한다.
_DEFAULT_COLUMNS = [
    ("confinas_code", "계정코드"),
    ("confinas_name", "계정명"),
    ("debit",         "차변합계"),
    ("credit",        "대변합계"),
    ("balance",       "잔액"),
    ("amount_krw",    "원화금액"),
    ("currency",      "통화"),
]


def _write_confinas_sheet(
    rows: list[dict],
    subsidiary_code: str,
    period: str,
    out_filepath: Path,
    template_columns: list[str] | None,
) -> None:
    out_filepath.parent.mkdir(parents=True, exist_ok=True)

    # template_columns가 있으면 그 순서를 사용, 없으면 기본 양식
    if template_columns:
        col_defs = [(c, c) for c in template_columns]
    else:
        col_defs = _DEFAULT_COLUMNS

    workbook  = xlsxwriter.Workbook(str(out_filepath))
    worksheet = workbook.add_worksheet("Confinas_Upload")

    # 스타일
    hdr_fmt = workbook.add_format({
        "bold": True, "bg_color": "#4472C4", "font_color": "#FFFFFF",
        "border": 1, "align": "center",
    })
    num_fmt = workbook.add_format({"num_format": "#,##0.00", "border": 1})
    txt_fmt = workbook.add_format({"border": 1})
    ttl_fmt = workbook.add_format({
        "bold": True, "bg_color": "#D9E1F2", "num_format": "#,##0.00", "border": 1,
    })

    # 타이틀
    worksheet.merge_range(
        0, 0, 0, len(col_defs) - 1,
        f"Confinas 업로드 데이터 — {subsidiary_code} {period}",
        workbook.add_format({"bold": True, "font_size": 13}),
    )

    # 헤더
    for col_idx, (_, header) in enumerate(col_defs):
        worksheet.write(1, col_idx, header, hdr_fmt)

    # 데이터
    for row_idx, row in enumerate(rows, start=2):
        for col_idx, (field, _) in enumerate(col_defs):
            val = row.get(field, "")
            if isinstance(val, float):
                worksheet.write_number(row_idx, col_idx, val, num_fmt)
            else:
                worksheet.write(row_idx, col_idx, val, txt_fmt)

    # 합계 행
    sum_row = len(rows) + 2
    worksheet.write(sum_row, 0, "합계", ttl_fmt)
    for col_idx, (field, _) in enumerate(col_defs):
        if col_idx == 0:
            continue
        if field in ("debit", "credit", "balance", "amount_krw"):
            total = sum(r.get(field, 0) for r in rows)
            worksheet.write_number(sum_row, col_idx, total, ttl_fmt)
        else:
            worksheet.write(sum_row, col_idx, "", ttl_fmt)

    # 열 너비 자동 조정
    worksheet.set_column(0, 0, 18)
    worksheet.set_column(1, 1, 30)
    worksheet.set_column(2, len(col_defs) - 1, 16)

    workbook.close()
