import click
from pathlib import Path
from rich.console import Console
from rich.table import Table
from rich import box

from config.settings import DATA_INPUT_DIR, SUBSIDIARY_CODES
from ingestion.excel_parser import parse_excel
from ingestion.data_normalizer import normalize
from ingestion.odata_client import OneCODataClient
from validation.engine import ValidationEngine
from models.validation_result import Severity
from utils.exceptions import OsstemBaseError

console = Console()


@click.group()
def cli():
    """Osstem AI — 해외법인 재무제표 자동 검증 시스템"""


# ═════════════════════════════════════════════════════════════════════════════
# recon — 3-System Reconciliation (현지회계 / 네트라 / Confinas)
# ═════════════════════════════════════════════════════════════════════════════

@cli.group("recon")
def recon():
    """현지회계 · 네트라 · Confinas 3-시스템 계정 대사 및 업로드"""


@recon.command("init-db")
def init_db():
    """MySQL에 필요한 테이블을 자동 생성합니다 (최초 1회 실행)."""
    from db.connection import get_engine, test_connection
    from db.models import Base

    if not test_connection():
        console.print("[red]X MySQL 연결 실패. 접속 정보를 확인하세요.[/red]")
        return

    try:
        Base.metadata.create_all(get_engine())
        console.print("[green]OK 테이블 생성 완료 (account_master / financial_local / financial_netra / upload_log)[/green]")
    except Exception as e:
        console.print(f"[red]X 테이블 생성 실패: {e}[/red]")


@recon.command("master-import")
@click.option("--file",  "filepath", required=True, help="계정 마스터 엑셀 파일 경로")
@click.option("--sheet", default="0", help="시트 번호 또는 이름 (기본: 0)")
def master_import(filepath, sheet):
    sheet = int(sheet) if sheet.isdigit() else sheet
    """계정과목 마스터(3-시스템 매핑)를 엑셀에서 DB로 임포트합니다.

    \b
    엑셀 필수 컬럼: subsidiary_code, local_code, netra_code, confinas_code
    (한글 컬럼명도 인식: 법인코드, 현지코드, 네트라코드, confinas코드 등)
    """
    from reconciliation.master_table import import_from_excel
    from db.connection import test_connection

    if not test_connection():
        console.print("[red]X MySQL 연결 실패. 접속 정보를 확인하세요.[/red]")
        return

    try:
        count = import_from_excel(filepath, sheet)
        console.print(f"[green]OK 계정 마스터 {count}건 임포트 완료[/green]")
    except Exception as e:
        console.print(f"[red]X 임포트 실패: {e}[/red]")


@recon.command("master-list")
@click.option("--corp", default=None, help="법인 코드 필터 (없으면 전체)")
def master_list(corp):
    """DB에 저장된 계정 마스터를 출력합니다."""
    from reconciliation.master_table import list_masters

    rows = list_masters(corp)
    if not rows:
        console.print("[yellow]! 마스터 데이터가 없습니다.[/yellow]")
        return

    table = Table(box=box.SIMPLE, show_header=True, header_style="bold cyan")
    table.add_column("법인", width=6)
    table.add_column("현지코드", width=12)
    table.add_column("현지계정명", width=20)
    table.add_column("네트라코드", width=12)
    table.add_column("Confinas코드", width=14)
    table.add_column("신계정", width=18)

    for r in rows:
        table.add_row(
            r["subsidiary"], r["local_code"] or "",
            r["local_name"] or "", r["netra_code"] or "",
            r["confinas_code"] or "", r["standard_code"] or "",
        )

    console.print(table)
    console.print(f"총 {len(rows)}건")


@recon.command("upload-local")
@click.option("--corp",   required=True, help="법인 코드 (예: UZ01)")
@click.option("--period", required=True, help="기간 (예: 2025-03)")
@click.option("--file",   "filepath", required=True, help="현지회계 엑셀 파일 경로")
@click.option("--sheet",  default=0, help="시트 번호 또는 이름 (기본: 0)")
def upload_local(corp, period, filepath, sheet):
    """현지회계프로그램 엑셀 데이터를 MySQL에 업로드합니다."""
    from reconciliation.uploader import upload_local as _upload

    console.print(f"[cyan]▶ {corp} {period} 현지회계 업로드 중...[/cyan]")
    result = _upload(corp, period, filepath, sheet)
    _print_upload_result(result)


@recon.command("upload-netra")
@click.option("--corp",   required=True, help="법인 코드 (예: UZ01)")
@click.option("--period", required=True, help="기간 (예: 2025-03)")
@click.option("--file",   "filepath", required=True, help="네트라 엑셀 파일 경로")
@click.option("--sheet",  default=0, help="시트 번호 또는 이름 (기본: 0)")
def upload_netra(corp, period, filepath, sheet):
    """네트라 엑셀 데이터를 MySQL에 업로드합니다."""
    from reconciliation.uploader import upload_netra as _upload

    console.print(f"[cyan]▶ {corp} {period} 네트라 업로드 중...[/cyan]")
    result = _upload(corp, period, filepath, sheet)
    _print_upload_result(result)


@recon.command("verify")
@click.option("--corp",   required=True, help="법인 코드 (예: UZ01)")
@click.option("--period", required=True, help="기간 (예: 2025-03)")
def verify(corp, period):
    """업로드 완료 여부를 검증합니다 (대차균형 · 행 수 · 누락계정)."""
    from reconciliation.verifier import verify as _verify

    result = _verify(corp, period)

    for system in ("local", "netra"):
        label = "현지회계" if system == "local" else "네트라"
        info  = result[system]

        if not info["uploaded"]:
            console.print(f"[red]X {label}: 업로드 데이터 없음[/red]")
            continue

        balanced_icon = "[green]OK[/green]" if info["is_balanced"] else "[red]X[/red]"
        console.print(
            f"{balanced_icon} [bold]{label}[/bold]  "
            f"행: {info['row_count']:,}  "
            f"차변: {info['total_debit']:,.0f}  "
            f"대변: {info['total_credit']:,.0f}  "
            f"대차균형: {'OK' if info['is_balanced'] else 'NG'}"
        )
        if info["missing_accounts"]:
            console.print(
                f"  [yellow]! 마스터 대비 누락 계정 {len(info['missing_accounts'])}건: "
                f"{', '.join(info['missing_accounts'][:5])}"
                + (" ..." if len(info['missing_accounts']) > 5 else "") + "[/yellow]"
            )

    # 업로드 이력
    if result["upload_log"]:
        console.print("\n[bold]업로드 이력 (최근 10건):[/bold]")
        log_table = Table(box=box.SIMPLE, show_header=True)
        log_table.add_column("시스템", width=10)
        log_table.add_column("상태", width=8)
        log_table.add_column("행수", width=8)
        log_table.add_column("대차균형", width=8)
        log_table.add_column("업로드일시", width=20)
        log_table.add_column("메시지")

        for log in result["upload_log"]:
            status_color = "green" if log["status"] == "success" else "red"
            balanced = "OK" if log["is_balanced"] else "NG" if log["is_balanced"] is False else "-"
            log_table.add_row(
                log["system"],
                f"[{status_color}]{log['status']}[/{status_color}]",
                str(log["row_count"] or ""),
                balanced,
                log["uploaded_at"],
                log["message"] or "",
            )
        console.print(log_table)


@recon.command("upload-netra-direct")
@click.option("--corp",   required=True, help="법인 코드 (예: UZ01)")
@click.option("--period", required=True, help="기간 (예: 2025-03)")
@click.option("--매출채권", "receivable", default=None, type=float, help="매출채권 금액")
@click.option("--선수금",   "advance",    default=None, type=float, help="선수금 금액")
@click.option("--원가",     "cogs",       default=None, type=float, help="원가 금액")
@click.option("--재고자산", "inventory",  default=None, type=float, help="재고자산 금액")
@click.option("--매출액",   "revenue",    default=None, type=float, help="매출액 금액")
@click.option("--currency", default="KRW", help="통화 (기본: KRW)")
@click.option("--rate",     default=1.0, type=float, help="환율 (기본: 1.0)")
def upload_netra_direct(corp, period, receivable, advance, cogs, inventory, revenue, currency, rate):
    """네트라 5개 항목 금액을 직접 입력하여 업로드합니다.

    \b
    예) python main.py recon upload-netra-direct --corp UZ01 --period 2025-03 \\
          --매출채권 1234567 --선수금 234567 --원가 3456789 \\
          --재고자산 5678901 --매출액 12345678 --currency UZS --rate 0.106
    """
    from reconciliation.uploader import upload_netra_direct as _upload

    data = {}
    if receivable is not None: data["매출채권"] = receivable
    if advance    is not None: data["선수금"]   = advance
    if cogs       is not None: data["원가"]     = cogs
    if inventory  is not None: data["재고자산"] = inventory
    if revenue    is not None: data["매출액"]   = revenue

    if not data:
        console.print("[red]X 최소 1개 항목 금액을 입력하세요.[/red]")
        return

    result = _upload(corp, period, data, currency, rate)
    if result["status"] == "success":
        console.print(f"[green]OK 네트라 직접입력 완료 ({result['row_count']}개 항목)[/green]")
    else:
        console.print(f"[red]X 실패: {result['message']}[/red]")


@recon.command("compare")
@click.option("--corp",     required=True, help="법인 코드 (예: UZ01)")
@click.option("--period",   required=True, help="기간 (예: 2025-03)")
@click.option("--detail",   is_flag=True,  help="각 항목별 구성 계정 상세 출력")
def compare(corp, period, detail):
    """현지회계(계정 집계) vs 네트라(5개 항목 합계)를 비교합니다."""
    from reconciliation.comparator import compare as _compare

    rows = _compare(corp, period)

    if not rows:
        console.print("[yellow]! 비교할 데이터가 없습니다.[/yellow]")
        console.print("  1) recon master-import — 계정 마스터에 네트라 항목 설정 확인")
        console.print("  2) recon upload-local   — 현지회계 데이터 업로드")
        console.print("  3) recon upload-netra   또는 upload-netra-direct — 네트라 데이터 입력")
        return

    flagged = sum(1 for r in rows if r["flagged"])
    console.print(
        f"\n[bold]현지회계(집계) vs 네트라 — {corp} {period}[/bold]  "
        f"[yellow]차이 항목: {flagged}/5[/yellow]\n"
    )

    table = Table(box=box.SIMPLE, show_header=True, header_style="bold")
    table.add_column("항목",     width=10)
    table.add_column("현지집계", width=18, justify="right")
    table.add_column("네트라",   width=18, justify="right")
    table.add_column("차이(원화)", width=18, justify="right")
    table.add_column("차이율",   width=8,  justify="right")
    table.add_column("구성계정수", width=8, justify="right")

    for r in rows:
        color = "red" if r["flagged"] else "green" if r["diff"] == 0 else "yellow"
        rate_str = f"{r['diff_rate']:.1%}"
        table.add_row(
            f"[bold]{r['category']}[/bold]",
            f"{r['local_total']:>18,.0f}",
            f"{r['netra_total']:>18,.0f}",
            f"[{color}]{r['diff']:>+18,.0f}[/{color}]",
            f"[{color}]{rate_str}[/{color}]",
            str(len(r["local_accounts"])),
        )

    console.print(table)

    # 상세 모드: 각 항목별 구성 계정 출력
    if detail:
        for r in rows:
            if not r["local_accounts"]:
                continue
            console.print(f"\n  [bold]{r['category']}[/bold] 구성 계정:")
            for acc in r["local_accounts"]:
                console.print(
                    f"    {acc['local_code']:<12} {acc['local_name']:<25} "
                    f"{acc['amount_krw']:>16,.0f} 원"
                )


@recon.command("list-local")
@click.option("--corp",   required=True, help="법인 코드 (예: UZ01)")
@click.option("--period", required=True, help="기간 (예: 2025-03)")
@click.option("--limit",  default=0, type=int, help="출력 행 수 제한 (기본: 전체)")
@click.option("--level",  default=None, type=click.Choice(["1","2","3"]),
              help="계정 레벨 필터 (1=주계정, 2=보조계정, 3=세부계정)")
@click.option("--excel",  "excel_path", default=None, help="엑셀 파일로도 저장")
def list_local(corp, period, limit, level, excel_path):
    """업로드된 현지회계 데이터를 차변/대변 구분해서 출력합니다."""
    from reconciliation.viewer import list_local as _list, export_local_excel

    level_int = int(level) if level else None
    data = _list(corp, period, limit, level=level_int)

    if not data["rows"]:
        console.print("[yellow]! 업로드된 데이터가 없습니다.[/yellow]")
        return

    is_ok = data["is_balanced"]
    balance_color = "green" if is_ok else "red"
    lv_label = f"  [dim]Lv{level} 필터[/dim]" if level else ""

    console.print(
        f"\n[bold]{corp} {period} 현지회계 시산표[/bold]{lv_label}  "
        f"총 {data['row_count']:,}행  "
        f"차변합: {data['total_debit']:,.0f}  "
        f"대변합: {data['total_credit']:,.0f}  "
        f"[{balance_color}]대차균형: {'OK' if is_ok else 'NG'}[/{balance_color}]\n"
    )

    _LV_STYLE = {1: "bold cyan", 2: "", 3: "dim"}

    table = Table(box=box.SIMPLE_HEAD, show_header=True, header_style="bold cyan")
    table.add_column("Lv",            width=4,  justify="center")
    table.add_column("계정코드",       width=14)
    table.add_column("계정명",         width=28)
    table.add_column("차변(Debit)",    width=16, justify="right")
    table.add_column("대변(Credit)",   width=16, justify="right")
    table.add_column("잔액(Balance)", width=16, justify="right")
    table.add_column("통화",           width=6)
    table.add_column("원화금액",       width=16, justify="right")

    for r in data["rows"]:
        lv = r["local_level"] or 2
        st = _LV_STYLE.get(lv, "")
        bal_color = "white" if r["balance"] >= 0 else "yellow"
        table.add_row(
            f"[{st}]Lv{lv}[/{st}]" if st else f"Lv{lv}",
            f"[{st}]{r['account_code']}[/{st}]" if st else r["account_code"],
            f"[{st}]{r['account_name']}[/{st}]" if st else r["account_name"],
            f"{r['debit']:>16,.0f}" if r["debit"] else "",
            f"{r['credit']:>16,.0f}" if r["credit"] else "",
            f"[{bal_color}]{r['balance']:>+16,.0f}[/{bal_color}]",
            r["currency"],
            f"{r['amount_krw']:>16,.0f}",
        )

    # 합계행
    table.add_row(
        "", "[bold]합  계[/bold]", "",
        f"[bold]{data['total_debit']:>16,.0f}[/bold]",
        f"[bold]{data['total_credit']:>16,.0f}[/bold]",
        f"[{balance_color}][bold]{data['total_debit'] - data['total_credit']:>+16,.0f}[/bold][/{balance_color}]",
        "", "",
    )

    console.print(table)

    if limit and data["row_count"] > limit:
        console.print(f"[dim]처음 {limit}행만 표시. --limit 0 으로 전체 출력[/dim]")

    if excel_path:
        from reconciliation.viewer import export_local_excel
        cnt = export_local_excel(corp, period, excel_path)
        console.print(f"[green]OK 엑셀 저장: {excel_path}  ({cnt}행)[/green]")


@recon.command("list-netra")
@click.option("--corp",   required=True, help="법인 코드 (예: UZ01)")
@click.option("--period", required=True, help="기간 (예: 2025-03)")
def list_netra(corp, period):
    """업로드된 네트라 5개 항목 데이터를 출력합니다."""
    from reconciliation.viewer import list_netra as _list

    data = _list(corp, period)

    if not data["rows"]:
        console.print("[yellow]! 네트라 데이터가 없습니다.[/yellow]")
        return

    console.print(f"\n[bold]{corp} {period} 네트라 데이터[/bold]\n")

    table = Table(box=box.SIMPLE_HEAD, show_header=True, header_style="bold cyan")
    table.add_column("항목",         width=12)
    table.add_column("금액(현지통화)", width=18, justify="right")
    table.add_column("통화",          width=6)
    table.add_column("환율",          width=10, justify="right")
    table.add_column("원화금액",      width=18, justify="right")

    total_krw = 0
    for r in data["rows"]:
        table.add_row(
            f"[bold]{r['category']}[/bold]",
            f"{r['amount']:>18,.0f}",
            r["currency"],
            f"{r['exchange_rate']:>10.4f}" if r["exchange_rate"] else "-",
            f"{r['amount_krw']:>18,.0f}",
        )
        total_krw += r["amount_krw"]

    table.add_row(
        "[bold]합  계[/bold]", "", "", "",
        f"[bold]{total_krw:>18,.0f}[/bold]",
    )

    console.print(table)


@recon.command("export-confinas")
@click.option("--corp",   required=True, help="법인 코드 (예: UZ01)")
@click.option("--period", required=True, help="기간 (예: 2025-03)")
@click.option("--out",    "out_filepath", required=True, help="출력 엑셀 파일 경로")
def export_confinas(corp, period, out_filepath):
    """현지회계 데이터를 Confinas 업로드용 엑셀로 변환합니다."""
    from reconciliation.confinas_exporter import export as _export

    console.print(f"[cyan]▶ {corp} {period} Confinas 엑셀 생성 중...[/cyan]")
    try:
        count = _export(corp, period, out_filepath)
        console.print(
            f"[green]OK Confinas 엑셀 생성 완료: {out_filepath}  ({count}개 계정)[/green]"
        )
    except Exception as e:
        console.print(f"[red]X 생성 실패: {e}[/red]")


def _print_upload_result(result: dict) -> None:
    if result["status"] == "success":
        is_bal  = result.get("is_balanced", True)
        bal_txt = "[green]대차균형 OK[/green]" if is_bal else "[red]대차균형 NG ![/red]"
        console.print(
            f"[green]OK 업로드 완료[/green]  "
            f"행: {result['row_count']:,}  "
            f"차변합: {result.get('total_debit', 0):,.0f}  "
            f"대변합: {result.get('total_credit', 0):,.0f}  "
            f"{bal_txt}"
        )
        console.print("[dim]  (확인) recon list-local --corp ... --period ...[/dim]")
    else:
        console.print(f"[red]X 업로드 실패: {result['message']}[/red]")


@cli.command("validate")
@click.option("--corp", default="all", help='법인 코드 (예: UZ01) 또는 "all"')
@click.option("--period", required=True, help='검증 기간 (예: 2025-03)')
@click.option("--source", default="excel", type=click.Choice(["excel", "1c"]),
              help="데이터 소스: excel(기본) 또는 1c(OData)")
@click.option("--sheet", default=0, help="엑셀 시트 번호 또는 이름 (excel 소스 전용)")
@click.option("--1c-url", "onec_url", default=None, help="1C 서버 URL (1c 소스 전용)")
@click.option("--1c-user", "onec_user", default=None, help="1C 사용자명")
@click.option("--1c-pass", "onec_pass", default=None, help="1C 비밀번호")
@click.option("--1c-org", "onec_org", default=None, help="1C 조직명 (없으면 전체)")
@click.option("--rate", default=1.0, type=float, help="원화 환산 환율 (1c 소스 전용)")
@click.option("--currency", default="UZS", help="원본 통화 코드 (기본: UZS)")
def validate(corp, period, source, sheet,
             onec_url, onec_user, onec_pass, onec_org, rate, currency):
    """재무 데이터를 읽어 검증 규칙을 실행하고 결과를 출력합니다.

    \b
    엑셀 사용 예:
      python main.py validate --corp UZ01 --period 2025-03

    1C OData 사용 예:
      python main.py validate --corp UZ01 --period 2025-03 \\
        --source 1c --1c-url http://192.168.1.10/UZ_Base \\
        --1c-user admin --1c-pass 1234 --rate 0.086
    """
    targets = SUBSIDIARY_CODES if corp == "all" else [corp.upper()]
    engine  = ValidationEngine()

    # 1C 클라이언트 초기화 (1c 소스 선택 시)
    onec_client = None
    if source == "1c":
        if not all([onec_url, onec_user, onec_pass]):
            console.print("[red]X --1c-url, --1c-user, --1c-pass 옵션이 필요합니다.[/red]")
            return
        onec_client = OneCODataClient(onec_url, onec_user, onec_pass)
        if not onec_client.test_connection():
            console.print(f"[red]X 1C 서버 연결 실패: {onec_url}[/red]")
            return
        console.print(f"[green]OK 1C 연결 성공: {onec_url}[/green]")

    for code in targets:
        console.print(f"\n[bold cyan]▶ {code}  {period}  [{source.upper()}][/bold cyan]")

        try:
            if source == "excel":
                pattern = list(DATA_INPUT_DIR.glob(f"{code}*{period}*.xlsx"))
                if not pattern:
                    console.print(f"[yellow]! 파일을 찾을 수 없습니다: {DATA_INPUT_DIR}/{code}*{period}*.xlsx[/yellow]")
                    continue
                tb = parse_excel(pattern[0], code, period, sheet_name=sheet)
                tb = normalize(tb)
            else:
                tb = onec_client.fetch_trial_balance(
                    subsidiary_code=code,
                    period=period,
                    org_name=onec_org,
                    exchange_rate=rate,
                    currency=currency,
                )

            result = engine.run(tb)

        except OsstemBaseError as e:
            console.print(f"[red]X 처리 실패: {e}[/red]")
            continue

        _print_result(result)


@cli.command("1c-orgs")
@click.option("--1c-url", "onec_url", required=True, help="1C 서버 URL")
@click.option("--1c-user", "onec_user", required=True, help="1C 사용자명")
@click.option("--1c-pass", "onec_pass", required=True, help="1C 비밀번호")
def list_orgs(onec_url, onec_user, onec_pass):
    """1C에 등록된 조직 목록을 출력합니다. (--1c-org 값 확인용)"""
    client = OneCODataClient(onec_url, onec_user, onec_pass)
    try:
        orgs = client.list_organizations()
        console.print("\n[bold]1C 조직 목록:[/bold]")
        for org in orgs:
            console.print(f"  • {org}")
    except OsstemBaseError as e:
        console.print(f"[red]X {e}[/red]")


def _print_result(result) -> None:
    summary = result.summary()

    if result.is_clean:
        console.print("[green]OK 검증 통과 — 이슈 없음[/green]")
        return

    # 요약 배너
    parts = []
    for sev, color in [(Severity.CRITICAL, "red"), (Severity.ERROR, "red"),
                       (Severity.WARNING, "yellow"), (Severity.INFO, "blue")]:
        if sev.value in summary:
            parts.append(f"[{color}]{sev.value}: {summary[sev.value]}[/{color}]")
    console.print("  " + "  |  ".join(parts))

    # 상세 테이블
    table = Table(box=box.SIMPLE, show_header=True, header_style="bold")
    table.add_column("규칙 ID", width=10)
    table.add_column("심각도", width=10)
    table.add_column("메시지")

    color_map = {
        Severity.CRITICAL: "red",
        Severity.ERROR:    "red",
        Severity.WARNING:  "yellow",
        Severity.INFO:     "blue",
    }

    for issue in result.issues:
        c = color_map.get(issue.severity, "white")
        table.add_row(
            f"[{c}]{issue.rule_id}[/{c}]",
            f"[{c}]{issue.severity.value}[/{c}]",
            issue.message,
        )

    console.print(table)


if __name__ == "__main__":
    cli()
