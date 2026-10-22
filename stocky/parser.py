"""Parser del reporte "LISTADO DE EXISTENCIAS" del SAI.

Acepta dos formatos y los detecta solo: el histórico con campos separados por
``|`` y el reporte en columnas. En ambos casos produce el mismo ``StockReport``.
"""

from __future__ import annotations

import re
from pathlib import Path
from typing import Iterable, Sequence

import openpyxl

from .models import StockReport, StockRow

_NUMBER = re.compile(r"-?\d+")
FIELDS = 12  # separadores "|" por línea; 13 campos al partir
_HEADER_MARKERS = (
    "NIVEL DE AGREGACION",
    "DESCRIPCION PRODUCTO",
    "UN. EMP",
    "E X I S T E N C I A S",
)


class ParseError(Exception):
    """El archivo no parece un listado de existencias utilizable."""


def _last_int(text: str) -> int:
    found = _NUMBER.findall(text)
    return int(found[-1]) if found else 0


def _is_digits(text: str) -> bool:
    return bool(text) and text.isdigit()


def _split(line: str) -> list[str] | None:
    if "|" not in line:
        return None
    fields = [part.strip() for part in line.split("|")]
    return fields if len(fields) == FIELDS + 1 else None


def _is_item(fields: list[str]) -> bool:
    return _is_digits(fields[1]) and _is_digits(fields[3])


def _is_continuation(fields: list[str]) -> bool:
    return not fields[1] and not fields[3] and _is_digits(fields[4])


def _item(fields: list[str]) -> StockRow:
    return StockRow(
        area=int(fields[1]),
        rubro=fields[2],
        plu=fields[3],
        barcode=fields[4],
        desc=fields[5],
        un_emp=_last_int(fields[6]),
        un_despa=_last_int(fields[7]),
        existencia=_last_int(fields[8]),
        disponible=_last_int(fields[9]),
        precio=_last_int(fields[10]),
        costo=_last_int(fields[11]),
    )


def _is_header_line(line: str) -> bool:
    if set(line.strip()) <= {"|", "-", " "}:
        return True
    if "--------" in line:
        return True
    return any(marker in line for marker in _HEADER_MARKERS)


_FECHA = re.compile(r"^\s*FECHA\s*:\s*(\d{2}/\d{2}/\d{4})")
_HORA = re.compile(r"^\s*HORA\s*:\s*(\d{1,2}:\d{2}:\d{2})")
_DEPENDENCIA = re.compile(r"^\s*DEPENDENCIA\s*:\s*(.+?)\s*$")


def _read_metadata(line: str, report: StockReport) -> None:
    if not report.fecha and (match := _FECHA.match(line)):
        report.fecha = match.group(1)
    elif not report.hora and (match := _HORA.match(line)):
        report.hora = match.group(1)
    elif not report.dependencia and (match := _DEPENDENCIA.match(line)):
        report.dependencia = match.group(1)


def _dedupe(rows: list[StockRow]) -> tuple[list[StockRow], int]:
    """Fusiona PLU repetidos; devuelve las filas únicas y los conflictos."""
    kept: dict[tuple[int, str], StockRow] = {}
    unique: list[StockRow] = []
    conflicts = 0
    for row in rows:
        key = (row.area, row.plu)
        previous = kept.get(key)
        if previous is None:
            kept[key] = row
            unique.append(row)
            continue
        same_data = (previous.desc, previous.disponible, previous.precio, previous.costo) == (
            row.desc,
            row.disponible,
            row.precio,
            row.costo,
        )
        if not same_data:
            conflicts += 1
            continue
        known = {previous.barcode, *previous.alt_barcodes}
        for code in (row.barcode, *row.alt_barcodes):
            if code and code not in known:
                previous.alt_barcodes.append(code)
                known.add(code)
    return unique, conflicts


def _looks_like_table(rows: Iterable[Sequence[object]]) -> bool:
    """Detecta el reporte en columnas por su fila de encabezado."""
    for row in rows:
        cells = {str(value).strip() for value in row if value is not None}
        if "PLU" in cells and any(cell.startswith("DESCRIPCION PRODUCTO") for cell in cells):
            return True
    return False


def parse_table_rows(rows: Iterable[Sequence[object]]) -> StockReport:
    """Lee una hoja en columnas (B=área, D=PLU, F=descripción, G=stock, H=precio)."""
    report = StockReport()
    for raw in rows:
        values = ["" if value is None else str(value).strip() for value in raw]
        if len(values) < 8:
            values.extend([""] * (8 - len(values)))
        _read_metadata("".join(values), report)
        area, rubro, plu, barcode, desc, disponible, precio = values[1:8]
        if not (area.isdigit() and 5 <= len(plu) <= 8 and plu.isdigit() and barcode and desc):
            continue
        report.rows.append(
            StockRow(
                area=int(area),
                rubro=rubro,
                plu=plu,
                barcode=barcode,
                desc=desc,
                un_emp=0,
                un_despa=0,
                existencia=0,
                disponible=_last_int(disponible),
                precio=_last_int(precio),
                costo=0,
            )
        )
    report.rows, report.conflicts = _dedupe(report.rows)
    if not report.rows:
        raise ParseError("El reporte no trae ítems")
    return report


def parse_rows(lines: Iterable[str]) -> StockReport:
    report = StockReport()
    for line in lines:
        if "|" not in line:
            _read_metadata(line, report)
            continue
        if _is_header_line(line):
            continue
        fields = _split(line)
        if fields is None:
            # Línea con pipes que no tiene la forma esperada: no se pierde en
            # silencio, se cuenta como desconocida.
            report.unknown_lines += 1
            continue
        if _is_item(fields):
            report.rows.append(_item(fields))
        elif _is_continuation(fields):
            if report.rows:
                report.rows[-1].alt_barcodes.append(fields[4])
            else:
                report.unknown_lines += 1
        else:
            report.unknown_lines += 1
    report.rows, report.conflicts = _dedupe(report.rows)
    if not report.rows:
        raise ParseError("El reporte no trae ítems")
    return report


def read_report(path: Path) -> StockReport:
    path = Path(path)
    workbook = openpyxl.load_workbook(path, data_only=True, read_only=True)
    try:
        sheet = workbook.active
        rows = list(sheet.iter_rows(values_only=True))
    finally:
        workbook.close()

    if _looks_like_table(rows):
        report = parse_table_rows(rows)
    else:
        lines = []
        for row in rows:
            cells = [str(cell) for cell in row if cell is not None]
            if cells:
                lines.append("".join(cells))
        if not any("|" in line for line in lines):
            raise ParseError("Formato no soportado: no parece un Listado de existencias")
        report = parse_rows(lines)

    report.source = path
    return report
