"""업로드된 재무 데이터 조회 — 차변/대변 구분 표시"""
from pathlib import Path

import xlsxwriter
from sqlalchemy import func, select

from db.connection import get_session
from db.models import FinancialLocal, FinancialNetra
from reconciliation.master_table import NETRA_CATEGORIES


def list_local(
    subsidiary_code: str,
    period: str,
    limit: int = 0,
    level: int | None = None,
) -> dict:
    """financial_local 테이블 조회.

    Returns:
        {
          "rows": [{"account_code", "account_name", "debit", "credit",
                    "balance", "currency", "exchange_rate", "amount_krw"}, ...],
          "total_debit":  float,
          "total_credit": float,
          "is_balanced":  bool,
          "row_count":    int,
        }
    """
    subsidiary_code = subsidiary_code.upper()
    with get_session() as session:
        conditions = [
            FinancialLocal.subsidiary_code == subsidiary_code,
            FinancialLocal.period == period,
        ]
        if level is not None:
            conditions.append(FinancialLocal.local_level == level)
        stmt = (
            select(FinancialLocal)
            .where(*conditions)
            .order_by(FinancialLocal.account_code)
        )
        if limit:
            stmt = stmt.limit(limit)
        rows = session.execute(stmt).scalars().all()

        agg = session.execute(
            select(
                func.sum(FinancialLocal.debit).label("d"),
                func.sum(FinancialLocal.credit).label("c"),
                func.count(FinancialLocal.id).label("n"),
            ).where(*conditions)
        ).one()

    total_debit  = float(agg.d or 0)
    total_credit = float(agg.c or 0)

    return {
        "rows": [
            {
                "account_code":  r.account_code,
                "account_name":  r.account_name or "",
                "local_level":   r.local_level,
                "debit":         float(r.debit  or 0),
                "credit":        float(r.credit or 0),
                "balance":       float(r.balance or 0),
                "currency":      r.currency or "",
                "exchange_rate": float(r.exchange_rate or 0),
                "amount_krw":    float(r.amount_krw or 0),
            }
            for r in rows
        ],
        "total_debit":  total_debit,
        "total_credit": total_credit,
        "is_balanced":  abs(total_debit - total_credit) < 1.0,
        "row_count":    int(agg.n or 0),
    }


def list_netra(subsidiary_code: str, period: str) -> dict:
    """financial_netra 테이블 조회 (5개 항목).

    Returns:
        {
          "rows": [{"category", "amount", "currency", "exchange_rate", "amount_krw"}, ...],
          "row_count": int,
        }
    """
    subsidiary_code = subsidiary_code.upper()
    with get_session() as session:
        rows = session.execute(
            select(FinancialNetra).where(
                FinancialNetra.subsidiary_code == subsidiary_code,
                FinancialNetra.period == period,
            )
        ).scalars().all()

    cat_order = {c: i for i, c in enumerate(NETRA_CATEGORIES)}
    sorted_rows = sorted(rows, key=lambda r: cat_order.get(r.category, 99))

    return {
        "rows": [
            {
                "category":      r.category,
                "amount":        float(r.amount or 0),
                "currency":      r.currency or "KRW",
                "exchange_rate": float(r.exchange_rate or 1),
                "amount_krw":    float(r.amount_krw or 0),
            }
            for r in sorted_rows
        ],
        "row_count": len(sorted_rows),
    }


def export_local_excel(
    subsidiary_code: str,
    period: str,
    out_filepath: str | Path,
) -> int:
    """현지회계 데이터를 차변/대변 구분 엑셀로 출력한다."""
    data = list_local(subsidiary_code, period)
    rows = data["rows"]
    out_filepath = Path(out_filepath)
    out_filepath.parent.mkdir(parents=True, exist_ok=True)

    wb  = xlsxwriter.Workbook(str(out_filepath))
    ws  = wb.add_worksheet(f"{subsidiary_code}_{period}_현지회계")

    # 서식
    hdr = wb.add_format({"bold": True, "bg_color": "#4472C4", "font_color": "#FFFFFF",
                          "border": 1, "align": "center"})
    num = wb.add_format({"num_format": "#,##0", "border": 1})
    txt = wb.add_format({"border": 1})
    ttl = wb.add_format({"bold": True, "bg_color": "#D9E1F2",
                          "num_format": "#,##0", "border": 1})
    red = wb.add_format({"bold": True, "font_color": "#FF0000",
                          "bg_color": "#FFE0E0", "num_format": "#,##0", "border": 1})

    # Lv별 행 배경색
    lv_bg = {
        1: wb.add_format({"border": 1, "bg_color": "#D9E1F2", "bold": True}),
        2: wb.add_format({"border": 1}),
        3: wb.add_format({"border": 1, "bg_color": "#F2F2F2", "italic": True}),
    }
    lv_num = {
        1: wb.add_format({"num_format": "#,##0", "border": 1, "bg_color": "#D9E1F2", "bold": True}),
        2: wb.add_format({"num_format": "#,##0", "border": 1}),
        3: wb.add_format({"num_format": "#,##0", "border": 1, "bg_color": "#F2F2F2", "italic": True}),
    }

    # 타이틀
    ws.merge_range("A1:I1",
                   f"현지회계 시산표 — {subsidiary_code} {period}  "
                   f"(대차균형: {'OK' if data['is_balanced'] else 'NG'})",
                   wb.add_format({"bold": True, "font_size": 12}))

    # 헤더
    headers = ["Lv", "계정코드", "계정명", "차변(Debit)", "대변(Credit)",
               "잔액(Balance)", "통화", "환율", "원화금액"]
    for ci, h in enumerate(headers):
        ws.write(1, ci, h, hdr)

    # 데이터
    for ri, r in enumerate(rows, start=2):
        lv = r["local_level"] or 2
        tf = lv_bg.get(lv, txt)
        nf = lv_num.get(lv, num)
        ws.write(ri, 0, f"Lv{lv}",           tf)
        ws.write(ri, 1, r["account_code"],    tf)
        ws.write(ri, 2, r["account_name"],    tf)
        ws.write(ri, 3, r["debit"],           nf)
        ws.write(ri, 4, r["credit"],          nf)
        ws.write(ri, 5, r["balance"],         nf)
        ws.write(ri, 6, r["currency"],        tf)
        ws.write(ri, 7, r["exchange_rate"],   nf)
        ws.write(ri, 8, r["amount_krw"],      nf)

    # 합계행
    sr = len(rows) + 2
    ws.write(sr, 0, "", ttl)
    ws.write(sr, 1, "합  계", ttl)
    ws.write(sr, 2, "", ttl)
    ws.write(sr, 3, data["total_debit"],  ttl)
    ws.write(sr, 4, data["total_credit"], ttl)
    diff = data["total_debit"] - data["total_credit"]
    ws.write(sr, 5, diff, ttl if abs(diff) < 1 else red)
    for ci in range(6, 9):
        ws.write(sr, ci, "", ttl)

    # 열 너비
    ws.set_column(0, 0, 6)   # Lv
    ws.set_column(1, 1, 16)  # 계정코드
    ws.set_column(2, 2, 30)  # 계정명
    ws.set_column(3, 5, 16)  # 차변/대변/잔액
    ws.set_column(6, 6, 8)   # 통화
    ws.set_column(7, 7, 10)  # 환율
    ws.set_column(8, 8, 16)  # 원화금액

    wb.close()
    return len(rows)
