"""업로드 완료 검증: 대차균형 · 행 수 · 필수 계정 존재 여부 확인"""
from sqlalchemy import func, select

from db.connection import get_session
from db.models import AccountMaster, FinancialLocal, FinancialNetra, UploadLog


def verify(subsidiary_code: str, period: str) -> dict:
    """현지회계 + 네트라 업로드 상태를 검증하고 결과를 반환한다.

    Returns:
        {
          "local":  {"uploaded": bool, "row_count": int, "total_debit": float,
                     "total_credit": float, "is_balanced": bool, "missing_accounts": list},
          "netra":  {...},
          "upload_log": [최근 이력 5건],
        }
    """
    subsidiary_code = subsidiary_code.upper()

    with get_session() as session:
        # ── 1. 업로드 이력 ────────────────────────────────────────────────────
        log_rows = session.execute(
            select(UploadLog)
            .where(
                UploadLog.subsidiary_code == subsidiary_code,
                UploadLog.period == period,
            )
            .order_by(UploadLog.uploaded_at.desc())
            .limit(10)
        ).scalars().all()

        # ── 2. 현지회계 통계 ──────────────────────────────────────────────────
        local_stats = session.execute(
            select(
                func.count(FinancialLocal.id).label("cnt"),
                func.sum(FinancialLocal.debit).label("debit"),
                func.sum(FinancialLocal.credit).label("credit"),
            ).where(
                FinancialLocal.subsidiary_code == subsidiary_code,
                FinancialLocal.period == period,
            )
        ).one()

        # ── 3. 네트라 통계 ────────────────────────────────────────────────────
        netra_stats = session.execute(
            select(
                func.count(FinancialNetra.id).label("cnt"),
                func.sum(FinancialNetra.debit).label("debit"),
                func.sum(FinancialNetra.credit).label("credit"),
            ).where(
                FinancialNetra.subsidiary_code == subsidiary_code,
                FinancialNetra.period == period,
            )
        ).one()

        # ── 4. 마스터 대비 누락 계정 탐지 ────────────────────────────────────
        master_local_codes = set(
            row.local_code
            for row in session.execute(
                select(AccountMaster.local_code).where(
                    AccountMaster.subsidiary_code == subsidiary_code,
                    AccountMaster.local_code.isnot(None),
                )
            ).all()
            if row.local_code
        )

        master_netra_codes = set(
            row.netra_code
            for row in session.execute(
                select(AccountMaster.netra_code).where(
                    AccountMaster.subsidiary_code == subsidiary_code,
                    AccountMaster.netra_code.isnot(None),
                )
            ).all()
            if row.netra_code
        )

        uploaded_local_codes = set(
            row.account_code
            for row in session.execute(
                select(FinancialLocal.account_code).where(
                    FinancialLocal.subsidiary_code == subsidiary_code,
                    FinancialLocal.period == period,
                )
            ).all()
        )

        uploaded_netra_codes = set(
            row.account_code
            for row in session.execute(
                select(FinancialNetra.account_code).where(
                    FinancialNetra.subsidiary_code == subsidiary_code,
                    FinancialNetra.period == period,
                )
            ).all()
        )

    # ── 집계 ──────────────────────────────────────────────────────────────────
    def _balanced(debit, credit) -> bool:
        d = float(debit or 0)
        c = float(credit or 0)
        return abs(d - c) < 1.0

    local_cnt    = local_stats.cnt or 0
    local_debit  = float(local_stats.debit or 0)
    local_credit = float(local_stats.credit or 0)

    netra_cnt    = netra_stats.cnt or 0
    netra_debit  = float(netra_stats.debit or 0)
    netra_credit = float(netra_stats.credit or 0)

    missing_local = sorted(master_local_codes - uploaded_local_codes)
    missing_netra = sorted(master_netra_codes - uploaded_netra_codes)

    return {
        "local": {
            "uploaded":        local_cnt > 0,
            "row_count":       local_cnt,
            "total_debit":     local_debit,
            "total_credit":    local_credit,
            "is_balanced":     _balanced(local_debit, local_credit),
            "missing_accounts": missing_local,
        },
        "netra": {
            "uploaded":        netra_cnt > 0,
            "row_count":       netra_cnt,
            "total_debit":     netra_debit,
            "total_credit":    netra_credit,
            "is_balanced":     _balanced(netra_debit, netra_credit),
            "missing_accounts": missing_netra,
        },
        "upload_log": [
            {
                "system":      r.system_name,
                "status":      r.status,
                "row_count":   r.row_count,
                "is_balanced": r.is_balanced,
                "uploaded_at": r.uploaded_at.strftime("%Y-%m-%d %H:%M:%S") if r.uploaded_at else "",
                "message":     r.message,
            }
            for r in log_rows
        ],
    }
