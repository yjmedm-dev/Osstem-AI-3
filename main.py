import click
from pathlib import Path
from rich.console import Console
from rich.table import Table
from rich import box

from config.settings import DATA_INPUT_DIR, SUBSIDIARY_CODES
from ingestion.excel_parser import parse_excel
from ingestion.schema_mapper import map_columns
from ingestion.data_normalizer import normalize
from validation.engine import ValidationEngine
from models.validation_result import Severity
from utils.exceptions import OsstemBaseError

console = Console()


@click.group()
def cli():
    """Osstem AI — 해외법인 재무제표 자동 검증 시스템"""


@cli.command("validate")
@click.option("--corp", default="all", help='법인 코드 (예: CN01) 또는 "all"')
@click.option("--period", required=True, help='검증 기간 (예: 2025-03)')
@click.option("--sheet", default=0, help="읽을 시트 번호 또는 이름 (기본: 0)")
def validate(corp: str, period: str, sheet):
    """엑셀 파일을 읽어 검증 규칙을 실행하고 결과를 출력합니다."""

    targets = SUBSIDIARY_CODES if corp == "all" else [corp.upper()]
    engine  = ValidationEngine()

    for code in targets:
        pattern = list(DATA_INPUT_DIR.glob(f"{code}*{period}*.xlsx"))
        if not pattern:
            console.print(f"[yellow]⚠ [{code}] {period} 파일을 찾을 수 없습니다.[/yellow]")
            continue

        file_path = pattern[0]
        console.print(f"\n[bold cyan]▶ {code}  {period}  {file_path.name}[/bold cyan]")

        try:
            tb = parse_excel(file_path, code, period, sheet_name=sheet)
            tb = normalize(tb)
            result = engine.run(tb)
        except OsstemBaseError as e:
            console.print(f"[red]✗ 처리 실패: {e}[/red]")
            continue

        _print_result(result)


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
