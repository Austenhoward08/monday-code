"""Typer-based CLI entry point for exporting Monday.com boards."""

from __future__ import annotations

import logging
from pathlib import Path
from typing import Optional

import typer
from rich.console import Console
from rich.traceback import install as install_rich_traceback

from .api import MondayClient
from .config import Settings
from .excel import ExcelExporter

install_rich_traceback(show_locals=False)

app = typer.Typer(
    add_completion=False,
    pretty_exceptions_show_locals=False,
    help="Export Monday.com board data into formatted Excel workbooks.",
)

console = Console()


@app.callback(invoke_without_command=True)
def main(
    board_id: int = typer.Option(..., "--board-id", "-b", help="Numeric ID of the Monday.com board."),
    output: Optional[Path] = typer.Option(
        None,
        "--output",
        "-o",
        help="Destination XLSX file. Defaults to '<board-name>.xlsx' in the current directory.",
    ),
    api_token: Optional[str] = typer.Option(
        None,
        "--api-token",
        envvar="MONDAY_API_TOKEN",
        help="Monday.com API token. Can also be set using the MONDAY_API_TOKEN environment variable.",
    ),
    include_subitems: bool = typer.Option(
        False,
        "--include-subitems",
        help="Include subitems (if present) for each item.",
    ),
    page_size: Optional[int] = typer.Option(
        None,
        "--page-size",
        min=1,
        max=1000,
        help="Override pagination size when retrieving items from Monday.com.",
    ),
    verbose: bool = typer.Option(
        False,
        "--verbose",
        "-v",
        help="Enable verbose logging output.",
    ),
) -> None:
    """Export a Monday.com board directly from the CLI."""
    if verbose:
        logging.basicConfig(level=logging.INFO, format="%(levelname)s %(name)s: %(message)s")

    settings_overrides = {}
    if page_size is not None:
        settings_overrides["page_size"] = page_size

    try:
        settings = Settings.from_env(api_token=api_token, overrides=settings_overrides)
    except RuntimeError as exc:
        console.print(f"[red]Configuration error:[/red] {exc}")
        raise typer.Exit(code=1) from exc

    try:
        with MondayClient(settings) as client:
            board = client.fetch_board(board_id=board_id, include_subitems=include_subitems)
    except Exception as exc:  # pragma: no cover - network exceptions are rare in tests
        console.print(f"[red]Failed to fetch board:[/red] {exc}")
        raise typer.Exit(code=1) from exc

    destination = output or _default_output_path(board.name)
    exporter = ExcelExporter()
    try:
        exported_path = exporter.export(board, destination)
    except Exception as exc:  # pragma: no cover
        console.print(f"[red]Failed to export to Excel:[/red] {exc}")
        raise typer.Exit(code=1) from exc

    console.print(
        f"[green]Successfully exported board[/green] [bold]{board.name}[/bold] "
        f"({len(board.items)} items) to [cyan]{exported_path}[/cyan]"
    )


def _default_output_path(board_name: str) -> Path:
    safe_name = "".join(ch if ch.isalnum() or ch in (" ", "-", "_") else "_" for ch in board_name)
    safe_name = "_".join(safe_name.strip().split())
    if not safe_name:
        safe_name = "monday_board"
    return Path(f"{safe_name}.xlsx")


__all__ = ["app"]
