"""
BSPL 엑셀 데이터를 RAG용 텍스트 청크로 변환한다.

청크 종류:
  1. 법인별 PL 요약 (핵심 계정 + 증감률)
  2. 법인별 BS 요약 (자산/부채/자본 3월말 잔액)
  3. 계정별 세부 (모든 PL 계정)
  4. 4개 법인 비교 요약
"""
from __future__ import annotations

import warnings
warnings.filterwarnings("ignore")

import pandas as pd
from dataclasses import dataclass, field
from pathlib import Path


@dataclass
class Chunk:
    chunk_id: str
    text: str
    metadata: dict = field(default_factory=dict)


# ── 설정 ──────────────────────────────────────────────────────────────
BSPL_FILES = {
    "러시아":     ("BSPL검토_러시아_2603.xlsx",     "RUB"),
    "우즈베키스탄": ("BSPL검토_우즈베키스탄_2603.xlsx", "UZS"),
    "우크라이나":  ("BSPL검토_우크라이나_2603.xlsx",  "UAH"),
    "카자흐스탄":  ("BSPL검토_카자흐스탄_2603.xlsx",  "KZT"),
}

PL_KEY_ITEMS = [
    ("PL01", "매출액"),
    ("PL02", "매출원가"),
    ("PL03", "매출총이익"),
    ("PL04", "판관비"),
    ("PL05", "영업이익"),
    ("PL06", "영업외수익"),
    ("PL07", "영업외비용"),
    ("PL08", "세전순이익"),
    ("PL09", "법인세"),
    ("PL10", "당기순이익"),
]

BS_KEY_ITEMS = [
    ("FP01", "자산총계"),
    ("FP02", "부채총계"),
    ("FP03", "자본총계"),
]

# PL 컬럼 인덱스 (헤더 없는 raw DataFrame 기준)
COL_PY_MONTH = 2   # 전년동기(월)
COL_PY_YTD   = 3   # 전년동기(누적)
COL_PY_END   = 4   # 전년말
COL_M1       = 5   # 1월
COL_M2       = 6   # 2월
COL_M3       = 7   # 3월
COL_CY_YTD   = 17  # 당기누적(1~3월)
COL_DELTA    = 18  # 증감


def _get_row(df: pd.DataFrame, code: str) -> pd.Series | None:
    rows = df[df.iloc[:, 0].astype(str).str.strip() == code]
    return rows.iloc[0] if len(rows) > 0 else None


def _pct(new: float, old: float) -> str:
    if old == 0:
        return "N/A"
    return f"{(new - old) / abs(old) * 100:+.1f}%"


def _fmt(v: float) -> str:
    return f"{v:,.0f}"


def _load_sheets(base_dir: Path, fname: str):
    path = base_dir / fname
    if not path.exists():
        # 루트 디렉토리에서도 탐색
        path = base_dir.parent / fname
    pl = pd.read_excel(str(path), sheet_name="PL", header=None)
    bs = pd.read_excel(str(path), sheet_name="BS", header=None)
    pl_data = pl[pl.iloc[:, 0].astype(str).str.match(r"^(PL|FP)\d+$")]
    bs_data = bs[bs.iloc[:, 0].astype(str).str.match(r"^(FP|FL)\d+$")]
    return pl_data, bs_data


# ── 청크 생성 함수들 ───────────────────────────────────────────────────

def _build_pl_summary_chunk(name: str, ccy: str, pl_data: pd.DataFrame, period: str) -> Chunk:
    lines = [f"[{name} 손익계산서 요약] 기간: {period}, 통화: {ccy}"]
    for code, nm in PL_KEY_ITEMS:
        r = _get_row(pl_data, code)
        if r is None:
            continue
        py_m  = float(r.iloc[COL_PY_MONTH] or 0)
        cy_m  = float(r.iloc[COL_M3] or 0)       # 3월
        cy_ytd = float(r.iloc[COL_CY_YTD] or 0)
        py_ytd = float(r.iloc[COL_PY_YTD] or 0)
        pct_str = _pct(cy_ytd, py_ytd)
        lines.append(
            f"  {code} {nm}: 당월(3월)={_fmt(cy_m)} {ccy}, "
            f"당기누적={_fmt(cy_ytd)} {ccy}, "
            f"전년동기(월)={_fmt(py_m)} {ccy}, "
            f"전년누적={_fmt(py_ytd)} {ccy}, "
            f"누적증감률={pct_str}"
        )
    # 마진 계산
    pl01 = _get_row(pl_data, "PL01")
    pl03 = _get_row(pl_data, "PL03")
    pl05 = _get_row(pl_data, "PL05")
    if pl01 is not None and pl03 is not None and pl05 is not None:
        rev = float(pl01.iloc[COL_M3] or 0)
        gp  = float(pl03.iloc[COL_M3] or 0)
        op  = float(pl05.iloc[COL_M3] or 0)
        if rev != 0:
            lines.append(f"  매출총이익률(3월): {gp/rev*100:.1f}%")
            lines.append(f"  영업이익률(3월): {op/rev*100:.1f}%")
    return Chunk(
        chunk_id=f"{name}_pl_summary_{period}",
        text="\n".join(lines),
        metadata={"corp": name, "ccy": ccy, "type": "pl_summary", "period": period},
    )


def _build_bs_summary_chunk(name: str, ccy: str, bs_data: pd.DataFrame, period: str) -> Chunk:
    lines = [f"[{name} 재무상태표 요약] 기간: {period}, 통화: {ccy}"]
    balances = {}
    for code, nm in BS_KEY_ITEMS:
        r = _get_row(bs_data, code)
        if r is None:
            continue
        py_end = float(r.iloc[COL_PY_END] or 0)
        m1     = float(r.iloc[COL_M1] or 0)
        m2     = float(r.iloc[COL_M2] or 0)
        m3     = float(r.iloc[COL_M3] or 0)
        cy_end = py_end + m1 + m2 + m3
        delta  = cy_end - py_end
        balances[code] = cy_end
        delta_str = f"+{_fmt(delta)}" if delta >= 0 else f"-{_fmt(abs(delta))}"
        lines.append(
            f"  {code} {nm}: 전년말={_fmt(py_end)} {ccy}, "
            f"3월말잔액={_fmt(cy_end)} {ccy}, "
            f"변동={delta_str} {ccy}"
        )
    # BS 균형 검증
    if len(balances) == 3:
        diff = balances["FP01"] - balances["FP02"] - balances["FP03"]
        status = "균형" if abs(diff) < 10 else f"불균형({_fmt(diff)})"
        lines.append(f"  BS 균형 검증: {status}")
    # 자본잠식 여부
    if "FP03" in balances:
        lines.append(f"  자본잠식 여부: {'자본잠식' if balances['FP03'] < 0 else '정상'}")
    return Chunk(
        chunk_id=f"{name}_bs_summary_{period}",
        text="\n".join(lines),
        metadata={"corp": name, "ccy": ccy, "type": "bs_summary", "period": period},
    )


def _build_pl_detail_chunks(name: str, ccy: str, pl_data: pd.DataFrame, period: str) -> list[Chunk]:
    """PL 계정별 세부 청크 (세부 코드 포함)"""
    chunks = []
    # PL01~PL10 각 그룹별 세부 계정 모음
    current_group = None
    lines = []

    def flush():
        nonlocal current_group, lines
        if current_group and lines:
            chunks.append(Chunk(
                chunk_id=f"{name}_pl_detail_{current_group}_{period}",
                text="\n".join(lines),
                metadata={"corp": name, "ccy": ccy, "type": "pl_detail",
                          "group": current_group, "period": period},
            ))
        current_group = None
        lines = []

    for _, row in pl_data.iterrows():
        code = str(row.iloc[0]).strip()
        name_col = str(row.iloc[1]) if pd.notna(row.iloc[1]) else ""
        cy_m   = float(row.iloc[COL_M3] or 0) if pd.notna(row.iloc[COL_M3]) else 0.0
        cy_ytd = float(row.iloc[COL_CY_YTD] or 0) if pd.notna(row.iloc[COL_CY_YTD]) else 0.0
        py_ytd = float(row.iloc[COL_PY_YTD] or 0) if pd.notna(row.iloc[COL_PY_YTD]) else 0.0

        # 최상위 그룹 변경
        import re
        if re.match(r"^PL\d{2}$", code):
            flush()
            current_group = code
            lines.append(f"[{name} {code} 세부] 기간: {period}, 통화: {ccy}")

        if current_group and cy_ytd != 0:
            pct_str = _pct(cy_ytd, py_ytd)
            lines.append(
                f"  {code} {name_col}: 3월={_fmt(cy_m)} {ccy}, "
                f"당기누적={_fmt(cy_ytd)} {ccy}, 전년누적={_fmt(py_ytd)} {ccy}, 증감={pct_str}"
            )

    flush()
    return chunks


def _build_comparison_chunk(all_data: dict, period: str) -> Chunk:
    """4개 법인 비교 요약 청크"""
    lines = [f"[4개 법인 손익 비교] 기간: {period}"]
    lines.append("(각 법인 로컬 통화 기준, 증감률은 전년동기누적 대비)")
    lines.append("")

    metrics = [
        ("PL01", "매출액", COL_M3, COL_PY_MONTH),
        ("PL03", "매출총이익", COL_M3, COL_PY_MONTH),
        ("PL05", "영업이익", COL_M3, COL_PY_MONTH),
        ("PL10", "당기순이익", COL_M3, COL_PY_MONTH),
    ]

    for code, nm, cy_col, py_col in metrics:
        lines.append(f"■ {nm}({code}) 당월(3월):")
        for corp_name, (pl_data, bs_data, ccy) in all_data.items():
            r = _get_row(pl_data, code)
            if r is None:
                continue
            cy  = float(r.iloc[cy_col] or 0)
            py  = float(r.iloc[py_col] or 0)
            pct = _pct(cy, py)
            lines.append(f"  {corp_name}({ccy}): {_fmt(cy)}, 전년비 {pct}")
        lines.append("")

    # 영업이익률 비교
    lines.append("■ 영업이익률(3월) 비교:")
    for corp_name, (pl_data, bs_data, ccy) in all_data.items():
        rev_r = _get_row(pl_data, "PL01")
        op_r  = _get_row(pl_data, "PL05")
        if rev_r is None or op_r is None:
            continue
        rev = float(rev_r.iloc[COL_M3] or 0)
        op  = float(op_r.iloc[COL_M3] or 0)
        margin = f"{op/rev*100:.1f}%" if rev != 0 else "N/A"
        lines.append(f"  {corp_name}: {margin}")

    # BS 자본잠식 비교
    lines.append("")
    lines.append("■ 재무상태 (3월말 자본총계):")
    for corp_name, (pl_data, bs_data, ccy) in all_data.items():
        r = _get_row(bs_data, "FP03")
        if r is None:
            continue
        py_end = float(r.iloc[COL_PY_END] or 0)
        cy_end = py_end + sum(float(r.iloc[i] or 0) for i in [COL_M1, COL_M2, COL_M3])
        status = "자본잠식" if cy_end < 0 else "정상"
        lines.append(f"  {corp_name}({ccy}): {_fmt(cy_end)} [{status}]")

    return Chunk(
        chunk_id=f"all_corps_comparison_{period}",
        text="\n".join(lines),
        metadata={"corp": "ALL", "type": "comparison", "period": period},
    )


# ── 진입점 ─────────────────────────────────────────────────────────────

def build_all_chunks(base_dir: str | Path, period: str = "2603") -> list[Chunk]:
    base_dir = Path(base_dir)
    chunks: list[Chunk] = []
    all_data: dict = {}

    for corp_name, (fname, ccy) in BSPL_FILES.items():
        pl_data, bs_data = _load_sheets(base_dir, fname)
        all_data[corp_name] = (pl_data, bs_data, ccy)

        chunks.append(_build_pl_summary_chunk(corp_name, ccy, pl_data, period))
        chunks.append(_build_bs_summary_chunk(corp_name, ccy, bs_data, period))
        chunks.extend(_build_pl_detail_chunks(corp_name, ccy, pl_data, period))

    chunks.append(_build_comparison_chunk(all_data, period))
    return chunks


if __name__ == "__main__":
    from pathlib import Path
    base = Path(__file__).resolve().parent.parent
    chunks = build_all_chunks(base)
    print(f"총 청크 수: {len(chunks)}")
    for c in chunks[:3]:
        print(f"\n--- {c.chunk_id} ---")
        print(c.text[:300])
