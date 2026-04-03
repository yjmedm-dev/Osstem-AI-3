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
            console.print("[red]✗ --1c-url, --1c-user, --1c-pass 옵션이 필요합니다.[/red]")
            return
        onec_client = OneCODataClient(onec_url, onec_user, onec_pass)
        if not onec_client.test_connection():
            console.print(f"[red]✗ 1C 서버 연결 실패: {onec_url}[/red]")
            return
        console.print(f"[green]✔ 1C 연결 성공: {onec_url}[/green]")

    for code in targets:
        console.print(f"\n[bold cyan]▶ {code}  {period}  [{source.upper()}][/bold cyan]")

        try:
            if source == "excel":
                pattern = list(DATA_INPUT_DIR.glob(f"{code}*{period}*.xlsx"))
                if not pattern:
                    console.print(f"[yellow]⚠ 파일을 찾을 수 없습니다: {DATA_INPUT_DIR}/{code}*{period}*.xlsx[/yellow]")
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
            console.print(f"[red]✗ 처리 실패: {e}[/red]")
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
        console.print(f"[red]✗ {e}[/red]")


def _print_result(result) -> None:
    summary = result.summary()

    if result.is_clean:
        console.print("[green]✔ 검증 통과 — 이슈 없음[/green]")
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
