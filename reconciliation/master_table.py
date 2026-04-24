"""계정과목 마스터 매핑 관리 (3개 시스템: 현지회계 / 네트라 / Confinas)"""
from pathlib import Path
from typing import Optional

import pandas as pd
from sqlalchemy import delete, select

from db.connection import get_session
from db.models import AccountMaster

# 엑셀 임포트 시 인식하는 컬럼명 → ORM 필드명 매핑
# 네트라 대사 가능 5개 항목 (category 컬럼에 저장되는 값)
NETRA_CATEGORIES = ["매출채권", "선수금", "원가", "재고자산", "매출액"]

_COLUMN_MAP = {
    "subsidiary_code":  "subsidiary_code",
    "법인코드":           "subsidiary_code",
    "local_code":       "local_code",
    "현지코드":           "local_code",
    "현지계정코드":        "local_code",
    "현지회계코드(1c)":   "local_code",
    "local_name":       "local_name",
    "현지계정명":          "local_name",
    "현지회계 계정명":    "local_name",
    "netra_category":   "netra_category",
    "네트라항목":          "netra_category",
    "네트라 항목":         "netra_category",
    "네트라대사항목":       "netra_category",
    "confinas_code":    "confinas_code",
    "confinas코드":     "confinas_code",
    "confinas 코드":    "confinas_code",
    "confinas_name":    "confinas_name",
    "confinas계정명":   "confinas_name",
    "confinas 계정명":  "confinas_name",
    "standard_code":    "standard_code",
    "신계정코드":          "standard_code",
    "신계정코드(fp/pl)":  "standard_code",
    "account_type":     "account_type",
    "계정유형":            "account_type",
}

_REQUIRED_FIELDS = ("subsidiary_code",)


def import_from_excel(filepath: str | Path, sheet: str | int = 0) -> int:
    """엑셀 파일에서 계정 마스터를 읽어 DB에 UPSERT 한다.

    Returns:
        삽입/갱신된 행 수
    """
    df = pd.read_excel(filepath, sheet_name=sheet, header=1, dtype=str).fillna("")

    # 컬럼명 정규화 (대소문자·공백 무시하여 매핑)
    normalized_map = {k.strip().lower(): v for k, v in _COLUMN_MAP.items()}
    df.columns = [c.strip() for c in df.columns]
    df = df.rename(columns=lambda c: normalized_map.get(c.strip().lower(), c))

    missing = [f for f in _REQUIRED_FIELDS if f not in df.columns]
    if missing:
        raise ValueError(f"마스터 엑셀에 필수 컬럼 없음: {missing}")

    rows = df.to_dict(orient="records")
    count = 0

    with get_session() as session:
        for row in rows:
            subsidiary = str(row.get("subsidiary_code", "")).strip().upper()
            local_code = str(row.get("local_code", "")).strip()
            if not subsidiary:
                continue

            # 기존 레코드 조회 (subsidiary_code + local_code 기준)
            stmt = select(AccountMaster).where(
                AccountMaster.subsidiary_code == subsidiary,
                AccountMaster.local_code == local_code,
            )
            obj = session.execute(stmt).scalar_one_or_none()

            if obj is None:
                obj = AccountMaster(
                    subsidiary_code=subsidiary,
                    local_code=local_code,
                )
                session.add(obj)

            netra_cat = str(row.get("netra_category", "")).strip()
            if netra_cat and netra_cat not in NETRA_CATEGORIES:
                netra_cat = ""  # 유효하지 않은 값은 무시

            obj.local_name      = str(row.get("local_name",    "")).strip() or None
            obj.netra_category  = netra_cat or None
            obj.confinas_code   = str(row.get("confinas_code", "")).strip() or None
            obj.confinas_name   = str(row.get("confinas_name", "")).strip() or None
            obj.standard_code   = str(row.get("standard_code", "")).strip() or None
            obj.account_type    = str(row.get("account_type",  "")).strip() or None
            count += 1

    return count


def get_mapping(subsidiary_code: str) -> dict:
    """법인별 전체 계정 매핑을 반환한다.

    Returns:
        {local_code: {netra_category, confinas_code, confinas_name, standard_code, account_type}}
    """
    with get_session() as session:
        rows = session.execute(
            select(AccountMaster).where(
                AccountMaster.subsidiary_code == subsidiary_code.upper()
            )
        ).scalars().all()

    return {
        r.local_code: {
            "local_name":      r.local_name,
            "netra_category":  r.netra_category,
            "confinas_code":   r.confinas_code,
            "confinas_name":   r.confinas_name,
            "standard_code":   r.standard_code,
            "account_type":    r.account_type,
        }
        for r in rows
        if r.local_code
    }


def get_confinas_mapping(subsidiary_code: str) -> dict:
    """confinas_code 기준 집계용 매핑을 반환한다.

    Returns:
        {local_code: confinas_code}
    """
    mapping = get_mapping(subsidiary_code)
    return {
        lc: info["confinas_code"]
        for lc, info in mapping.items()
        if info.get("confinas_code")
    }


def list_masters(subsidiary_code: Optional[str] = None) -> list[dict]:
    """마스터 테이블 전체(또는 특정 법인) 조회"""
    with get_session() as session:
        stmt = select(AccountMaster)
        if subsidiary_code:
            stmt = stmt.where(
                AccountMaster.subsidiary_code == subsidiary_code.upper()
            )
        rows = session.execute(stmt).scalars().all()

    return [
        {
            "id":              r.id,
            "subsidiary":      r.subsidiary_code,
            "local_code":      r.local_code,
            "local_name":      r.local_name,
            "netra_category":  r.netra_category,
            "confinas_code":   r.confinas_code,
            "confinas_name":   r.confinas_name,
            "standard_code":   r.standard_code,
            "account_type":    r.account_type,
        }
        for r in rows
    ]


def delete_master(subsidiary_code: str) -> int:
    """특정 법인의 마스터 레코드 전체 삭제. 재임포트 전 초기화 용도."""
    with get_session() as session:
        result = session.execute(
            delete(AccountMaster).where(
                AccountMaster.subsidiary_code == subsidiary_code.upper()
            )
        )
        return result.rowcount
