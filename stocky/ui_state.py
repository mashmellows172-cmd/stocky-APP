"""Selección de áreas y textos de estado del panel."""

from __future__ import annotations

from .models import StockReport

DEFAULT_AREAS = [8, 9, 29, 46, 84, 99]


def initial_selection(disponibles: list[int], last: list[int]) -> list[int]:
    """Áreas marcadas al abrir: las últimas usadas o las seis de siempre."""
    available = set(disponibles)
    chosen = [area for area in last if area in available]
    if not chosen:
        chosen = [area for area in DEFAULT_AREAS if area in available]
    if not chosen:
        chosen = sorted(available)
    return chosen


def _miles(number: int) -> str:
    return f"{number:,}".replace(",", ".")


def status_text(report: StockReport, selection: list[int]) -> str:
    parts = [
        f"{_miles(len(report.rows))} ítems",
        f"{len(selection)} {'área' if len(selection) == 1 else 'áreas'}",
    ]
    if report.unknown_lines:
        singular = report.unknown_lines == 1
        parts.append(
            f"{report.unknown_lines} línea no reconocida" if singular else f"{report.unknown_lines} líneas no reconocidas"
        )
    if report.conflicts:
        singular = report.conflicts == 1
        parts.append(
            f"{report.conflicts} duplicado conflictivo" if singular else f"{report.conflicts} duplicados conflictivos"
        )
    return " · ".join(parts)


def resolve_font(family: str, available: list[str]) -> str:
    return family if family in available else "Segoe UI"
