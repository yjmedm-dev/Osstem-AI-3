"""현지회계 vs 네트라 5개 항목 집계 비교"""
from sqlalchemy import select

from db.connection import get_session
from db.models import AccountMaster, FinancialLocal, FinancialNetra
from reconciliation.master_table import NETRA_CATEGORIES

_DIFF_RATE_THRESHOLD   = 0.01       # 1% 초과 시 하이라이트
_DIFF_AMOUNT_THRESHOLD = 1_000_000  # 100만원 초과 시 하이라이트


def compare(subsidiary_code: str, period: str) -> list[dict]:
    """로컬 계정을 netra_category로 집계하고 네트라 합계와 비교한다.

    Returns:
        [
          {
            "category":        str,    # 매출채권 / 선수금 / 원가 / 재고자산 / 매출액
            "local_total":     float,  # 현지회계 해당 계정 합산 잔액(원화)
            "netra_total":     float,  # 네트라 보고 금액(원화)
            "diff":            float,  # local - netra
            "diff_rate":       float,  # |diff| / |local| (local=0이면 1.0)
            "flagged":         bool,
            "local_accounts":  list[dict],  # 집계에 포함된 현지 계정 목록
          },
          ...
        ]
    """
    subsidiary_code = subsidiary_code.upper()

    with get_session() as session:
        master_rows = session.execute(
            select(AccountMaster).where(
                AccountMaster.subsidiary_code == subsidiary_code,
                AccountMaster.netra_category.isnot(None),
            )
        ).scalars().all()

        local_rows = session.execute(
            select(FinancialLocal).where(
                FinancialLocal.subsidiary_code == subsidiary_code,
                FinancialLocal.period == period,
            )
        ).scalars().all()

        netra_rows = session.execute(
            select(FinancialNetra).where(
                FinancialNetra.subsidiary_code == subsidiary_code,
                FinancialNetra.period == period,
            )
        ).scalars().all()

    # 현지회계: account_code → 잔액/계정명
    local_data = {
        r.account_code: {
            "account_name": r.account_name or "",
            "balance":      float(r.balance or 0),
            "amount_krw":   float(r.amount_krw or 0),
        }
        for r in local_rows
    }

    # 마스터: local_code → netra_category
    cat_map: dict[str, str] = {
        m.local_code: m.netra_category
        for m in master_rows
        if m.local_code
    }

    # 네트라: category → 원화금액
    netra_totals: dict[str, float] = {
        r.category: float(r.amount_krw or 0)
        for r in netra_rows
    }

    # category별 현지 계정 집계
    category_local: dict[str, list] = {cat: [] for cat in NETRA_CATEGORIES}
    for lc, cat in cat_map.items():
        if cat not in category_local:
            continue
        d = local_data.get(lc)
        if d is None:
            continue
        category_local[cat].append({
            "local_code":  lc,
            "local_name":  d["account_name"],
            "balance":     d["balance"],
            "amount_krw":  d["amount_krw"],
        })

    results = []
    for cat in NETRA_CATEGORIES:
        accounts   = category_local.get(cat, [])
        local_tot  = sum(a["amount_krw"] for a in accounts)
        netra_tot  = netra_totals.get(cat, 0.0)
        diff       = local_tot - netra_tot

        if local_tot != 0:
            diff_rate = abs(diff) / abs(local_tot)
        elif netra_tot != 0:
            diff_rate = 1.0
        else:
            diff_rate = 0.0

        flagged = (
            diff_rate > _DIFF_RATE_THRESHOLD
            or abs(diff) > _DIFF_AMOUNT_THRESHOLD
        )

        results.append({
            "category":       cat,
            "local_total":    local_tot,
            "netra_total":    netra_tot,
            "diff":           diff,
            "diff_rate":      diff_rate,
            "flagged":        flagged,
            "local_accounts": sorted(accounts, key=lambda x: abs(x["amount_krw"]), reverse=True),
        })

    return results


def compare_detail(subsidiary_code: str, period: str, category: str) -> list[dict]:
    """특정 category에 속하는 현지 계정 상세 내역을 반환한다."""
    all_results = compare(subsidiary_code, period)
    for r in all_results:
        if r["category"] == category:
            return r["local_accounts"]
    return []
