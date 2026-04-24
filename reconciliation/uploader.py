"""현지회계프로그램 / 네트라 엑셀 → MySQL 업로드 파이프라인

네트라 업로드 방식:
  네트라는 개별 계정코드 없이 5개 항목(매출채권/선수금/원가/재고자산/매출액) 합계만 제공.
  업로드 방식 2가지:
  1. 엑셀 파일: 첫 열=항목명, 둘째 열=금액 (다른 열은 통화/환율/원화금액)
  2. 직접 입력: upload_netra_direct(corp, period, {항목: 금액}) 호출
"""
from datetime import datetime
from pathlib import Path
from typing import Literal

import pandas as pd
from sqlalchemy import delete, select

from db.connection import get_session
from db.models import FinancialLocal, FinancialNetra, UploadLog
from reconciliation.master_table import NETRA_CATEGORIES

# 네트라 엑셀에서 항목 열/금액 열로 인식하는 헤더명
_NETRA_ITEM_COL  = ["항목", "item", "category", "구분", "계정"]
_NETRA_AMT_COL   = ["금액", "amount", "잔액", "balance"]
_NETRA_KRW_COL   = ["원화금액", "amount_krw", "원화환산액", "원화잔액"]
_NETRA_CURR_COL  = ["통화", "currency"]
_NETRA_RATE_COL  = ["환율", "exchange_rate"]

# 현지회계 엑셀 컬럼 매핑 (법인별 상이할 수 있음)
_LOCAL_COLUMN_MAP = {
    "account_code":  "account_code",
    "account_name":  "account_name",
    "debit":         "debit",
    "credit":        "credit",
    "balance":       "balance",
    "currency":      "currency",
    "exchange_rate": "exchange_rate",
    "amount_krw":    "amount_krw",
    # 한글
    "계정코드":       "account_code",
    "계정과목코드":    "account_code",
    "계정명":         "account_name",
    "계정과목명":      "account_name",
    "차변":           "debit",
    "대변":           "credit",
    "잔액":           "balance",
    "통화":           "currency",
    "환율":           "exchange_rate",
    "원화금액":        "amount_krw",
    # 러시아어/우즈벡 헤더 (1C 직접 출력)
    "Счет":          "account_code",
    "Субконто":      "account_name",
    "Дебет":         "debit",
    "Кредит":        "credit",
}


def _read_excel(filepath: str | Path, col_map: dict, sheet: str | int = 0) -> pd.DataFrame:
    df = pd.read_excel(filepath, sheet_name=sheet, dtype=str).fillna("")
    df.columns = [c.strip() for c in df.columns]
    df = df.rename(columns=col_map)
    return df


def _to_float(val) -> float:
    try:
        return float(str(val).replace(",", "").strip())
    except (ValueError, TypeError):
        return 0.0


def _detect_level(code: str) -> int:
    """계정코드 → 레벨 분류.
    Lv3: 소수점 포함 (예: 0120.2, 4410.1)
    Lv1: 끝 2자리가 '00' (예: 0100, 1000, 5000)
    Lv2: 그 외 4자리 (예: 0120, 5010, 4010)
    """
    s = str(code).strip()
    if '.' in s:
        return 3
    digits = s.replace(' ', '')
    if len(digits) >= 2 and digits[-2:] == '00':
        return 1
    return 2


def _build_rows(df: pd.DataFrame, subsidiary_code: str, period: str, model_cls):
    """DataFrame → ORM 객체 리스트 변환"""
    rows = []
    for _, r in df.iterrows():
        acc_code = str(r.get("account_code", "")).strip()
        if not acc_code:
            continue

        debit  = _to_float(r.get("debit",  0))
        credit = _to_float(r.get("credit", 0))

        # balance 컬럼이 있으면 우선 사용, 없으면 차변-대변 계산
        if "balance" in r and str(r["balance"]).strip():
            balance = _to_float(r.get("balance"))
        else:
            balance = debit - credit

        # 원화 금액: amount_krw 컬럼 있으면 사용, 없으면 balance 사용
        if "amount_krw" in r and str(r["amount_krw"]).strip():
            amount_krw = _to_float(r.get("amount_krw"))
        else:
            amount_krw = balance

        extra = {}
        if hasattr(model_cls, "local_level"):
            extra["local_level"] = _detect_level(acc_code)

        rows.append(model_cls(
            subsidiary_code=subsidiary_code,
            period=period,
            account_code=acc_code,
            account_name=str(r.get("account_name", "")).strip() or None,
            debit=debit,
            credit=credit,
            balance=balance,
            currency=str(r.get("currency", "")).strip() or None,
            exchange_rate=_to_float(r.get("exchange_rate", 0)) or None,
            amount_krw=amount_krw,
            uploaded_at=datetime.utcnow(),
            **extra,
        ))
    return rows


def _upload(
    system: Literal["local", "netra"],
    filepath: str | Path,
    subsidiary_code: str,
    period: str,
    col_map: dict,
    model_cls,
    sheet: str | int = 0,
) -> dict:
    """공통 업로드 로직"""
    subsidiary_code = subsidiary_code.upper()
    filepath = Path(filepath)

    try:
        df = _read_excel(filepath, col_map, sheet)
        orm_rows = _build_rows(df, subsidiary_code, period, model_cls)

        if not orm_rows:
            return {"status": "error", "message": "유효한 데이터 행이 없습니다."}

        total_debit  = sum(float(r.debit or 0)  for r in orm_rows)
        total_credit = sum(float(r.credit or 0) for r in orm_rows)
        is_balanced  = abs(total_debit - total_credit) < 1.0

        with get_session() as session:
            # 기존 데이터 삭제 (동일 법인·기간 재업로드 지원)
            session.execute(
                delete(model_cls).where(
                    model_cls.subsidiary_code == subsidiary_code,
                    model_cls.period == period,
                )
            )
            session.add_all(orm_rows)

            # 업로드 이력 기록
            session.add(UploadLog(
                system_name=system,
                subsidiary_code=subsidiary_code,
                period=period,
                status="success",
                row_count=len(orm_rows),
                total_debit=total_debit,
                total_credit=total_credit,
                is_balanced=is_balanced,
                message=f"파일: {filepath.name}",
            ))

        return {
            "status": "success",
            "row_count": len(orm_rows),
            "total_debit": total_debit,
            "total_credit": total_credit,
            "is_balanced": is_balanced,
        }

    except Exception as e:
        with get_session() as session:
            session.add(UploadLog(
                system_name=system,
                subsidiary_code=subsidiary_code,
                period=period,
                status="error",
                message=str(e),
            ))
        return {"status": "error", "message": str(e)}


def _parse_1c_tb(filepath: str | Path, sheet: str | int = 0) -> pd.DataFrame:
    """1C:Enterprise 잔액시산표 형식 파싱.

    파일 구조:
      Row 0~7: 타이틀/헤더 (스킵)
      Col 0 : "계정코드, 계정명" (쉼표 구분)
      Col 1 : 통화
      Col 2 : 기초잔액 차변  Col 3: 기초잔액 대변
      Col 4 : 당기 차변      Col 5: 당기 대변
      Col 7 : 기말잔액 차변  Col 8: 기말잔액 대변
    """
    import re
    df_raw = pd.read_excel(filepath, sheet_name=sheet, header=None, dtype=str).fillna("")

    records = []
    for _, row in df_raw.iterrows():
        cell0 = str(row.iloc[0]).strip()
        if not cell0 or cell0 == "nan":
            continue
        if "," not in cell0:
            continue

        code_part = cell0.split(",")[0].strip()
        name_part = cell0.split(",", 1)[1].strip()

        # 숫자 + 소수점으로만 이루어진 코드만 취급
        if not re.match(r"^\d+(\.\d+)?$", code_part):
            continue

        # 통화 컬럼 처리
        currency = str(row.iloc[1]).strip() if len(row) > 1 else ""
        # нат. = натуральный (수량 단위) — 금액 행 아님, 스킵
        if currency.startswith("нат"):
            continue
        # БУ = Бухгалтерский учет (회계장부 보고통화) → UZS로 대체
        if not currency or currency == "nan" or currency == "БУ":
            currency = "UZS"

        debit  = _to_float(row.iloc[7]) if len(row) > 7 else 0.0
        credit = _to_float(row.iloc[8]) if len(row) > 8 else 0.0

        records.append({
            "account_code":  code_part,
            "account_name":  name_part,
            "debit":         debit,
            "credit":        credit,
            "balance":       debit - credit,
            "currency":      currency,
            "exchange_rate": "",
            "amount_krw":    debit - credit,  # 환율 미적용, UZS 원화 그대로
        })

    return pd.DataFrame(records)


def _is_1c_format(filepath: str | Path, sheet: str | int = 0) -> bool:
    """파일 첫 컬럼이 '코드, 이름' 패턴이면 1C 형식으로 판단."""
    import re
    try:
        df = pd.read_excel(filepath, sheet_name=sheet, header=None, dtype=str, nrows=15).fillna("")
        for _, row in df.iterrows():
            cell = str(row.iloc[0]).strip()
            if re.match(r"^\d+(\.\d+)?,", cell):
                return True
    except Exception:
        pass
    return False


def upload_local(
    subsidiary_code: str,
    period: str,
    filepath: str | Path,
    sheet: str | int = 0,
) -> dict:
    """현지회계프로그램 엑셀 → financial_local 테이블 업로드
    1C:Enterprise 잔액시산표 형식 자동 감지 후 파싱.
    """
    filepath = Path(filepath)
    subsidiary_code = subsidiary_code.upper()

    try:
        if _is_1c_format(filepath, sheet):
            df = _parse_1c_tb(filepath, sheet)
        else:
            df = _read_excel(filepath, _LOCAL_COLUMN_MAP, sheet)

        orm_rows = _build_rows(df, subsidiary_code, period, FinancialLocal)

        if not orm_rows:
            return {"status": "error", "message": "유효한 데이터 행이 없습니다."}

        total_debit  = sum(float(r.debit  or 0) for r in orm_rows)
        total_credit = sum(float(r.credit or 0) for r in orm_rows)
        is_balanced  = abs(total_debit - total_credit) < 1.0

        with get_session() as session:
            session.execute(
                delete(FinancialLocal).where(
                    FinancialLocal.subsidiary_code == subsidiary_code,
                    FinancialLocal.period == period,
                )
            )
            session.add_all(orm_rows)
            session.add(UploadLog(
                system_name="local",
                subsidiary_code=subsidiary_code,
                period=period,
                status="success",
                row_count=len(orm_rows),
                total_debit=total_debit,
                total_credit=total_credit,
                is_balanced=is_balanced,
                message=f"파일: {filepath.name}",
            ))

        return {
            "status": "success",
            "row_count": len(orm_rows),
            "total_debit": total_debit,
            "total_credit": total_credit,
            "is_balanced": is_balanced,
        }

    except Exception as e:
        with get_session() as session:
            session.add(UploadLog(
                system_name="local",
                subsidiary_code=subsidiary_code,
                period=period,
                status="error",
                message=str(e),
            ))
        return {"status": "error", "message": str(e)}


# ─────────────────────────────────────────────────────────────────────────────
# 네트라 — 5개 항목 합계 업로드
# ─────────────────────────────────────────────────────────────────────────────

def _find_col(df: pd.DataFrame, candidates: list[str]) -> str | None:
    """대소문자·공백 무시해서 후보 헤더 중 첫 번째 일치하는 열명 반환"""
    normalized = {c.strip().lower(): c for c in df.columns}
    for cand in candidates:
        if cand.strip().lower() in normalized:
            return normalized[cand.strip().lower()]
    return None


def upload_netra(
    subsidiary_code: str,
    period: str,
    filepath: str | Path,
    sheet: str | int = 0,
) -> dict:
    """네트라 엑셀 → financial_netra 테이블 업로드 (5개 항목 합계 방식)

    엑셀 형식 (첫 열: 항목명, 둘째 열~: 금액/통화/환율/원화금액):
    ┌──────────┬───────────┬──────┬──────┬───────────┐
    │ 항목     │ 금액      │ 통화 │ 환율 │ 원화금액  │
    ├──────────┼───────────┼──────┼──────┼───────────┤
    │ 매출채권 │ 1,234,567 │ UZS  │ 0.11 │ 135,802   │
    │ 선수금   │   234,567 │ UZS  │ 0.11 │  25,802   │
    │ 원가     │ 3,456,789 │ UZS  │ 0.11 │ 380,247   │
    │ 재고자산 │ 5,678,901 │ UZS  │ 0.11 │ 624,679   │
    │ 매출액   │12,345,678 │ UZS  │ 0.11 │1,358,025  │
    └──────────┴───────────┴──────┴──────┴───────────┘
    """
    subsidiary_code = subsidiary_code.upper()
    filepath = Path(filepath)

    try:
        df = pd.read_excel(filepath, sheet_name=sheet, dtype=str).fillna("")
        df.columns = [c.strip() for c in df.columns]

        item_col = _find_col(df, _NETRA_ITEM_COL)
        amt_col  = _find_col(df, _NETRA_AMT_COL)
        krw_col  = _find_col(df, _NETRA_KRW_COL)
        curr_col = _find_col(df, _NETRA_CURR_COL)
        rate_col = _find_col(df, _NETRA_RATE_COL)

        # 첫 열을 항목, 둘째 열을 금액으로 fallback
        if item_col is None:
            item_col = df.columns[0]
        if amt_col is None and len(df.columns) > 1:
            amt_col = df.columns[1]

        orm_rows = []
        for _, row in df.iterrows():
            category = str(row[item_col]).strip()
            if category not in NETRA_CATEGORIES:
                continue

            amount     = _to_float(row.get(amt_col, 0) if amt_col else 0)
            amount_krw = _to_float(row.get(krw_col, 0) if krw_col else amount)
            currency   = str(row.get(curr_col, "")).strip() if curr_col else None
            rate       = _to_float(row.get(rate_col, 0)) if rate_col else None

            orm_rows.append(FinancialNetra(
                subsidiary_code=subsidiary_code,
                period=period,
                category=category,
                amount=amount,
                currency=currency or None,
                exchange_rate=rate or None,
                amount_krw=amount_krw,
                uploaded_at=datetime.utcnow(),
            ))

        if not orm_rows:
            return {
                "status": "error",
                "message": (
                    f"항목명이 일치하지 않습니다. "
                    f"첫 열에 {NETRA_CATEGORIES} 중 하나가 있어야 합니다."
                ),
            }

        with get_session() as session:
            session.execute(
                delete(FinancialNetra).where(
                    FinancialNetra.subsidiary_code == subsidiary_code,
                    FinancialNetra.period == period,
                )
            )
            session.add_all(orm_rows)
            session.add(UploadLog(
                system_name="netra",
                subsidiary_code=subsidiary_code,
                period=period,
                status="success",
                row_count=len(orm_rows),
                message=f"파일: {filepath.name}",
            ))

        return {"status": "success", "row_count": len(orm_rows),
                "categories": [r.category for r in orm_rows]}

    except Exception as e:
        with get_session() as session:
            session.add(UploadLog(
                system_name="netra",
                subsidiary_code=subsidiary_code,
                period=period,
                status="error",
                message=str(e),
            ))
        return {"status": "error", "message": str(e)}


def upload_netra_direct(
    subsidiary_code: str,
    period: str,
    data: dict,
    currency: str = "KRW",
    exchange_rate: float | None = None,
) -> dict:
    """네트라 데이터를 직접 dict로 입력한다.

    Args:
        data: {항목명: 금액} — 예) {"매출채권": 1234567, "매출액": 5000000}
        currency: 원본 통화 (기본 KRW — 원화이면 amount_krw = amount)
        exchange_rate: 환율 (KRW이면 1.0)
    """
    subsidiary_code = subsidiary_code.upper()
    rate = exchange_rate or 1.0

    try:
        orm_rows = []
        for cat, amt in data.items():
            cat = cat.strip()
            if cat not in NETRA_CATEGORIES:
                continue
            amount = float(amt)
            orm_rows.append(FinancialNetra(
                subsidiary_code=subsidiary_code,
                period=period,
                category=cat,
                amount=amount,
                currency=currency,
                exchange_rate=rate,
                amount_krw=amount * rate,
                uploaded_at=datetime.utcnow(),
            ))

        if not orm_rows:
            return {"status": "error", "message": "유효한 항목이 없습니다."}

        with get_session() as session:
            session.execute(
                delete(FinancialNetra).where(
                    FinancialNetra.subsidiary_code == subsidiary_code,
                    FinancialNetra.period == period,
                )
            )
            session.add_all(orm_rows)
            session.add(UploadLog(
                system_name="netra",
                subsidiary_code=subsidiary_code,
                period=period,
                status="success",
                row_count=len(orm_rows),
                message="직접입력",
            ))

        return {"status": "success", "row_count": len(orm_rows)}

    except Exception as e:
        return {"status": "error", "message": str(e)}


# ─────────────────────────────────────────────────────────────────────────────
# 네트라 소스 파일 3종(AR / SalesList / SBS) → financial_netra 업로드
# ─────────────────────────────────────────────────────────────────────────────

def upload_netra_from_sources(
    subsidiary_code: str,
    period: str,
    ar_filepath: str | Path,
    sales_filepath: str | Path,
    sbs_filepath: str | Path,
    currency: str = "UZS",
) -> dict:
    """AR / Sales List / SBS 파일에서 네트라 5개 항목을 집계해 업로드한다.

    집계 방식:
      매출채권  = AR BALANCE > 0 합계
      선수금    = AR BALANCE < 0 절댓값 합계
      매출액    = Sales List NET AMT 합계  (STEP1별 세부 저장)
      원가      = SBS col13 합계           (STEP1별 세부 저장)
      재고자산  = SBS col18 합계           (STEP1별 세부 저장)

    네트라 비교는 category 합계 기준이며,
    STEP1은 financial_netra.step1 에 세부 참고용으로 저장된다.
    현지회계는 STEP1 구분 없이 category 합계로만 비교한다.
    """
    subsidiary_code = subsidiary_code.upper()

    try:
        rows: list[FinancialNetra] = []

        # ── 1. 매출채권 / 선수금 (AR) ─────────────────────────────────────
        ar_df = pd.read_excel(ar_filepath, sheet_name=0, dtype=str).fillna("")
        ar_df.columns = [c.strip() for c in ar_df.columns]

        bal_col = next((c for c in ar_df.columns if c.upper() == "BALANCE"), None)
        if bal_col is None:
            return {"status": "error", "message": "AR 파일에 BALANCE 컬럼이 없습니다."}

        ar_df["_bal"] = ar_df[bal_col].apply(_to_float)
        receivable = ar_df[ar_df["_bal"] > 0]["_bal"].sum()
        advance    = abs(ar_df[ar_df["_bal"] < 0]["_bal"].sum())

        rows.append(FinancialNetra(
            subsidiary_code=subsidiary_code, period=period,
            category="매출채권", step1=None,
            amount=receivable, currency=currency,
            exchange_rate=None, amount_krw=receivable,
            uploaded_at=datetime.utcnow(),
        ))
        rows.append(FinancialNetra(
            subsidiary_code=subsidiary_code, period=period,
            category="선수금", step1=None,
            amount=advance, currency=currency,
            exchange_rate=None, amount_krw=advance,
            uploaded_at=datetime.utcnow(),
        ))

        # ── 2. 매출액 (Sales List — STEP1별) ─────────────────────────────
        sl_df = pd.read_excel(sales_filepath, sheet_name=0, dtype=str).fillna("")
        sl_df.columns = [c.strip() for c in sl_df.columns]
        sl_df["_net"] = sl_df["NET AMT"].apply(_to_float)

        step1_col = "STEP 1" if "STEP 1" in sl_df.columns else "STEP1"
        for step1, grp in sl_df.groupby(step1_col):
            amt = grp["_net"].sum()
            rows.append(FinancialNetra(
                subsidiary_code=subsidiary_code, period=period,
                category="매출액", step1=str(step1),
                amount=amt, currency=currency,
                exchange_rate=None, amount_krw=amt,
                uploaded_at=datetime.utcnow(),
            ))

        # ── 3. 원가 / 재고자산 (SBS — STEP1별) ───────────────────────────
        sbs_raw = pd.read_excel(sbs_filepath, sheet_name=0, header=None, dtype=str).fillna("")
        sbs_df  = sbs_raw.iloc[1:].copy()
        sbs_df.columns = range(len(sbs_df.columns))
        sbs_df["_cogs"] = sbs_df[13].apply(_to_float)
        sbs_df["_inv"]  = sbs_df[18].apply(_to_float)

        for step1, grp in sbs_df.groupby(19):
            cogs = grp["_cogs"].sum()
            inv  = grp["_inv"].sum()
            rows.append(FinancialNetra(
                subsidiary_code=subsidiary_code, period=period,
                category="원가", step1=str(step1),
                amount=cogs, currency=currency,
                exchange_rate=None, amount_krw=cogs,
                uploaded_at=datetime.utcnow(),
            ))
            rows.append(FinancialNetra(
                subsidiary_code=subsidiary_code, period=period,
                category="재고자산", step1=str(step1),
                amount=inv, currency=currency,
                exchange_rate=None, amount_krw=inv,
                uploaded_at=datetime.utcnow(),
            ))

        if not rows:
            return {"status": "error", "message": "집계된 데이터가 없습니다."}

        with get_session() as session:
            session.execute(
                delete(FinancialNetra).where(
                    FinancialNetra.subsidiary_code == subsidiary_code,
                    FinancialNetra.period == period,
                )
            )
            session.add_all(rows)
            session.add(UploadLog(
                system_name="netra",
                subsidiary_code=subsidiary_code,
                period=period,
                status="success",
                row_count=len(rows),
                message=f"AR:{Path(ar_filepath).name} / SL:{Path(sales_filepath).name} / SBS:{Path(sbs_filepath).name}",
            ))

        # category별 합계 요약
        from collections import defaultdict
        cat_totals: dict[str, float] = defaultdict(float)
        for r in rows:
            cat_totals[r.category] += float(r.amount or 0)

        return {
            "status":     "success",
            "row_count":  len(rows),
            "categories": dict(cat_totals),
        }

    except Exception as e:
        with get_session() as session:
            session.add(UploadLog(
                system_name="netra",
                subsidiary_code=subsidiary_code,
                period=period,
                status="error",
                message=str(e),
            ))
        return {"status": "error", "message": str(e)}
