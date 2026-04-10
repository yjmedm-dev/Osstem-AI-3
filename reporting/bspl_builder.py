# -*- coding: utf-8 -*-
"""
BS/PL 빌더 (우즈베키스탄 UZ01 연습용)

데이터 소스:
  - RAW 시트  : 기말잔액(BS) + 당기 오버턴(PL 매출/원가/영업외)
  - PL세부 시트: 판관비 항목별 상세 금액

출력:
  - BS (재무상태표) : FP01 자산 / FP02 부채 / FP03 자본
  - PL (손익계산서) : 매출 → 원가 → 판관비 → 영업외 → 법인세
"""

from __future__ import annotations
from pathlib import Path
from dataclasses import dataclass, field
import pandas as pd

from ingestion.schema_mapper import _UZ01_ACCOUNT_MAP

# ─────────────────────────────────────────────────────────────────
# 연결계정 계층 정의 (IFRS Report 기준)
# ─────────────────────────────────────────────────────────────────
# (code, name, parent_prefix)  — 리프 노드만 정의, 합계는 자동 계산
BS_HIERARCHY = [
    # ── 자산 ─────────────────────────────────────────────
    ("FP01",          "자산",                     None),
    ("FP01-01",       "유동자산",                  "FP01"),
    ("FP01-01-01",    "당좌자산",                  "FP01-01"),
    ("FP01-01-01-0010", "현금및현금성자산",          "FP01-01-01"),
    ("FP01-01-01-0070", "매출채권",                "FP01-01-01"),
    ("FP01-01-01-0080", "대손충당금(매출채권)",      "FP01-01-01"),
    ("FP01-01-01-0100", "미수금",                  "FP01-01-01"),
    ("FP01-01-01-0140", "단기대여금",               "FP01-01-01"),
    ("FP01-01-01-0170", "선급금",                  "FP01-01-01"),
    ("FP01-01-01-0190", "선급비용",                "FP01-01-01"),
    ("FP01-01-01-0200", "선급법인세",               "FP01-01-01"),
    ("FP01-01-01-0210", "선급부가세",               "FP01-01-01"),
    ("FP01-01-02",    "재고자산",                  "FP01-01"),
    ("FP01-01-02-0010-02", "본사제품매입분",         "FP01-01-02"),
    ("FP01-01-02-0110",    "저장품",               "FP01-01-02"),
    ("FP01-01-02-0120",    "미착품",               "FP01-01-02"),
    ("FP01-02",       "비유동자산",                "FP01"),
    ("FP01-02-02",    "유형자산",                  "FP01-02"),
    ("FP01-02-02-0110", "차량운반구",               "FP01-02-02"),
    ("FP01-02-02-0120", "감가상각충당금(차량운반구)", "FP01-02-02"),
    ("FP01-02-02-0160", "집기비품",                "FP01-02-02"),
    ("FP01-02-02-0170", "감가상각충당금(집기비품)",  "FP01-02-02"),
    ("FP01-02-03",    "무형자산",                  "FP01-02"),
    ("FP01-02-03-0050", "기타의무형자산",            "FP01-02-03"),
    ("FP01-02-04",    "기타비유동자산",             "FP01-02"),
    ("FP01-02-04-0030", "보증금",                  "FP01-02-04"),
    # ── 부채 ─────────────────────────────────────────────
    ("FP02",          "부채",                     None),
    ("FP02-01",       "유동부채",                  "FP02"),
    ("FP02-01-01-0010", "매입채무",                "FP02-01"),
    ("FP02-01-01-0110", "미지급금",                "FP02-01"),
    ("FP02-01-01-0120", "선수금",                  "FP02-01"),
    ("FP02-01-01-0150", "미지급비용",               "FP02-01"),
    ("FP02-01-01-0180", "부가세예수금",              "FP02-01"),
    # ── 자본 ─────────────────────────────────────────────
    ("FP03",          "자본",                     None),
    ("FP03-01",       "납입자본",                  "FP03"),
    ("FP03-01-01-0010", "보통주자본금",             "FP03-01"),
    ("FP03-03",       "자본조정",                  "FP03"),
    ("FP03-03-01-0070", "기타자본조정",              "FP03-03"),
    ("FP03-05",       "이익잉여금",                "FP03"),
    ("FP03-05-01-0060", "미처분이익잉여금_전기이월",  "FP03-05"),
    ("FP03-05-01-0070", "미처분이익잉여금_당기순손익","FP03-05"),
]

PL_HIERARCHY = [
    ("PL01",      "매출액",          None),
    ("PL01-01-0020", "본사제품매출액", "PL01"),
    ("PL02",      "매출원가",         None),
    ("PL02-01-0020", "본사제품매출원가","PL02"),
    ("PL03",      "매출총이익",        None),   # 계산값
    ("PL04",      "판매비와관리비",     None),
    ("PL04-00-0010", "급여",          "PL04"),
    ("PL04-00-0020", "상여금",         "PL04"),
    ("PL04-00-0050", "복리후생비",      "PL04"),
    ("PL04-00-0070", "여비교통비",      "PL04"),
    ("PL04-00-0090", "소모품비",        "PL04"),
    ("PL04-00-0110", "운반비",          "PL04"),
    ("PL04-00-0120", "통신비",          "PL04"),
    ("PL04-00-0130", "수도광열비",       "PL04"),
    ("PL04-00-0140", "임차료",          "PL04"),
    ("PL04-00-0150", "보험료",          "PL04"),
    ("PL04-00-0170", "차량유지비",       "PL04"),
    ("PL04-00-0180", "감가상각비",       "PL04"),
    ("PL04-00-0200", "무형자산상각비",    "PL04"),
    ("PL04-00-0210", "광고선전비",       "PL04"),
    ("PL04-00-0230", "세금과공과",       "PL04"),
    ("PL04-00-0240", "지급수수료",       "PL04"),
    ("PL05",      "영업이익",           None),  # 계산값
    ("PL06",      "영업외수익",          None),
    ("PL06-00-0050", "외화환산이익",      "PL06"),
    ("PL06-00-0280", "잡이익",           "PL06"),
    ("PL07",      "영업외비용",          None),
    ("PL07-00-0040", "외화환산손실",      "PL07"),
    ("PL07-00-0260", "잡손실",           "PL07"),
    ("PL08",      "법인세차감전순이익",    None),  # 계산값
    ("PL09",      "법인세비용",           None),
    ("PL09-00-0010", "법인세비용",        "PL09"),
    ("PL10",      "당기순손익",           None),  # 계산값
]

# 자동 계산되는 합계 코드
_CALC_CODES = {
    "PL03": lambda d: d.get("PL01", 0) - d.get("PL02", 0),
    "PL05": lambda d: d.get("PL03", 0) - d.get("PL04", 0),
    "PL08": lambda d: d.get("PL05", 0) + d.get("PL06", 0) - d.get("PL07", 0),
    "PL10": lambda d: d.get("PL08", 0) - d.get("PL09", 0),
}


# ─────────────────────────────────────────────────────────────────
# 파싱 함수
# ─────────────────────────────────────────────────────────────────

def _parse_raw_bs(xl: pd.ExcelFile) -> dict[str, float]:
    """RAW 시트 기말잔액 → {신계정코드: 순잔액(차변-대변)} 집계."""
    raw = pd.read_excel(xl, sheet_name="RAW", header=None)

    result: dict[str, float] = {}
    for _, row in raw.iterrows():
        acc_full = str(row.iloc[1]) if pd.notna(row.iloc[1]) else ""
        if not acc_full or acc_full == "nan":
            continue

        uz_code = acc_full.split(",")[0].strip()
        std_code = _UZ01_ACCOUNT_MAP.get(uz_code) or _UZ01_ACCOUNT_MAP.get(acc_full.strip())
        if not std_code:
            continue

        debit  = float(row.iloc[7]) if pd.notna(row.iloc[7]) else 0.0
        credit = float(row.iloc[8]) if pd.notna(row.iloc[8]) else 0.0
        result[std_code] = result.get(std_code, 0.0) + (debit - credit)

    return result


def _parse_raw_pl(xl: pd.ExcelFile) -> dict[str, float]:
    """RAW 시트 오버턴(당기) → {신계정코드: 금액} 집계.

    1C 월마감 구조상 PL 계정은 "원래 거래" 와 "9910 대체(마감)" 분개가
    양쪽에 동일 금액으로 쌓여 col5 == col6 이 된다.
    따라서 순액(col6-col5)이 0이 되므로, 단측(single-side) 오버턴을 사용한다.
      - 수익 계정 : col6(대변 오버턴) = 실제 수익액
      - 비용/차감 계정 : col5(차변 오버턴) = 실제 비용액
    반품(9040.1)은 매출에서 차감하므로 sign=-1.
    """
    raw = pd.read_excel(xl, sheet_name="RAW", header=None)
    result: dict[str, float] = {}

    # (신계정코드, 사용할 컬럼, 부호)
    #   컬럼 5 = 오버턴 차변(Д), 컬럼 6 = 오버턴 대변(К)
    pl_map: dict[str, tuple[str, int, float]] = {
        "9020.1": ("PL01-01-0020", 6, +1.0),   # 매출: 대변 오버턴
        "9040.1": ("PL01-01-0020", 5, -1.0),   # 반품: 차변 오버턴 (매출 차감)
        "9120.1": ("PL02-01-0020", 5, +1.0),   # 매출원가: 차변 오버턴
        "9540":   ("PL06-00-0050", 6, +1.0),   # 외화환산이익: 대변 오버턴
        "9541":   ("PL06-00-0050", 6, +1.0),
        "9620":   ("PL07-00-0040", 5, +1.0),   # 외화환산손실: 차변 오버턴
        "9690":   ("PL07-00-0260", 5, +1.0),   # 잡손실: 차변 오버턴
        "9810":   ("PL09-00-0010", 5, +1.0),   # 법인세: 차변 오버턴
        "9820":   ("PL09-00-0010", 5, +1.0),
    }

    for _, row in raw.iterrows():
        acc_full = str(row.iloc[1]) if pd.notna(row.iloc[1]) else ""
        uz_code  = acc_full.split(",")[0].strip()
        if uz_code not in pl_map:
            continue

        std_code, col_idx, sign = pl_map[uz_code]
        raw_val = row.iloc[col_idx]
        val = float(raw_val) if pd.notna(raw_val) else 0.0
        result[std_code] = result.get(std_code, 0.0) + sign * val

    return result


def _parse_pl_detail(xl: pd.ExcelFile) -> dict[str, float]:
    """PL세부 시트 → {신계정코드: 금액} 집계 (판관비 상세)."""
    df = pd.read_excel(xl, sheet_name="PL세부", header=0)
    result: dict[str, float] = {}

    for _, row in df.iterrows():
        name = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ""
        amount = row.iloc[1]
        if not name or name == "nan" or not pd.notna(amount):
            continue

        std_code = _UZ01_ACCOUNT_MAP.get(name)
        if not std_code:
            continue

        result[std_code] = result.get(std_code, 0.0) + float(amount)

    return result


# ─────────────────────────────────────────────────────────────────
# 집계 엔진
# ─────────────────────────────────────────────────────────────────

def _aggregate(hierarchy: list, leaf_values: dict[str, float]) -> dict[str, float]:
    """계층 구조를 따라 합계를 자동 계산한다."""
    totals: dict[str, float] = {}

    # 1단계: 리프 값 초기화
    for code, _, _ in hierarchy:
        totals[code] = leaf_values.get(code, 0.0)

    # 2단계: 하위 → 상위로 합산 (역순 순회)
    for code, _, parent in reversed(hierarchy):
        if parent and totals.get(code, 0.0) != 0.0:
            totals[parent] = totals.get(parent, 0.0) + totals[code]

    # 3단계: 자동 계산 항목
    for calc_code, fn in _CALC_CODES.items():
        totals[calc_code] = fn(totals)

    return totals


# ─────────────────────────────────────────────────────────────────
# 출력 함수
# ─────────────────────────────────────────────────────────────────

def _depth(code: str) -> int:
    return code.count("-")


def print_bs(totals: dict[str, float]) -> None:
    """재무상태표 출력."""
    print("=" * 65)
    print(f"{'재무상태표 (BS)':^65}")
    print(f"{'(단위: UZS)':^65}")
    print("=" * 65)

    for code, name, _ in BS_HIERARCHY:
        val = totals.get(code, 0.0)
        depth = _depth(code)
        indent = "  " * depth
        is_leaf = depth >= 3

        if is_leaf:
            print(f"  {indent}{name:<35} {val:>18,.0f}")
        else:
            # 합계 행 굵게 표시 (구분선)
            marker = "─" * (50 - depth * 2)
            print(f"\n  {indent}【{name}】")
            if depth >= 1:
                print(f"  {indent}{'':35} {val:>18,.0f}")

    # 검증: 자산 = 부채 + 자본
    asset  = totals.get("FP01", 0.0)
    liab   = abs(totals.get("FP02", 0.0))
    equity = abs(totals.get("FP03", 0.0))
    print()
    print("─" * 65)
    print(f"  {'자산 합계':<35} {asset:>18,.0f}")
    print(f"  {'부채 합계':<35} {liab:>18,.0f}")
    print(f"  {'자본 합계':<35} {equity:>18,.0f}")
    print(f"  {'부채 + 자본':<35} {liab+equity:>18,.0f}")
    diff = abs(asset) - (liab + equity)
    balance_flag = "[균형]" if abs(diff) < 1 else "[불균형]"
    print(f"  {'차이(자산 - 부채-자본)':<35} {diff:>18,.0f}  {balance_flag}")


def print_pl(totals: dict[str, float]) -> None:
    """손익계산서 출력."""
    print()
    print("=" * 65)
    print(f"{'손익계산서 (PL)':^65}")
    print(f"{'(단위: UZS)':^65}")
    print("=" * 65)

    calc_labels = {"PL03": "매출총이익", "PL05": "영업이익",
                   "PL08": "법인세차감전순이익", "PL10": "당기순손익"}

    for code, name, parent in PL_HIERARCHY:
        val = totals.get(code, 0.0)
        depth = _depth(code) if parent else 0
        indent = "  " * depth

        if code in calc_labels:
            print()
            print(f"  {'─'*60}")
            label = calc_labels[code]
            sign  = "▼ 손실" if val < 0 else "▲ 이익"
            print(f"  {indent}{label:<35} {val:>18,.0f}  {sign}")
        elif parent is None:
            print(f"\n  【{name}】")
            if val != 0.0:
                print(f"  {'':35} {val:>18,.0f}")
        else:
            print(f"  {indent}{name:<35} {val:>18,.0f}")


# ─────────────────────────────────────────────────────────────────
# 메인 진입점
# ─────────────────────────────────────────────────────────────────

def build_bspl(file_path: Path, period: str) -> None:
    """BS/PL 빌드 및 출력."""
    xl = pd.ExcelFile(file_path)

    # 데이터 수집
    bs_leaf  = _parse_raw_bs(xl)
    pl_raw   = _parse_raw_pl(xl)
    pl_detail = _parse_pl_detail(xl)

    # PL: RAW + PL세부 합산 (판관비는 PL세부 우선)
    pl_leaf: dict[str, float] = {}
    for k, v in pl_raw.items():
        pl_leaf[k] = pl_leaf.get(k, 0.0) + v
    for k, v in pl_detail.items():
        pl_leaf[k] = pl_leaf.get(k, 0.0) + v

    # 집계
    bs_totals = _aggregate(BS_HIERARCHY, bs_leaf)
    pl_totals = _aggregate(PL_HIERARCHY, pl_leaf)

    # 출력
    print(f"\n법인: UZ01 (우즈베키스탄)  기간: {period}\n")
    print_bs(bs_totals)
    print_pl(pl_totals)


if __name__ == "__main__":
    build_bspl(
        Path(__file__).parent.parent / "우즈벡 마감자료_2602.xlsx",
        period="2026-02",
    )
