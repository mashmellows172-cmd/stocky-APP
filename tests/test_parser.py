from pathlib import Path

import pytest

from stocky.parser import parse_rows, read_report

FIXTURE = Path(__file__).parent / "fixtures" / "reporte_recorte.txt"

ITEM_26726760 = (
    " |  1   |7881 MEDIANA                |26726760  |00200026726766 |"
    " AMERICAN TOURISTER MALETA FRANKFURT. |(UNI)   -1|(UNI)   -1|       -1 |       -1 |"
    "     59990 |     -49994|"
)
ITEM_26726760_CON_CODIGO_2 = (
    " |  1   |7881 MEDIANA                |26726760  |00200026726799 |"
    " AMERICAN TOURISTER MALETA FRANKFURT. |(UNI)   -1|(UNI)   -1|       -1 |       -1 |"
    "     59990 |     -49994|"
)
ITEM_26726760_STOCK_5 = (
    " |  1   |7881 MEDIANA                |26726760  |00200026726766 |"
    " AMERICAN TOURISTER MALETA FRANKFURT. |(UNI)    5|(UNI)    5|        5 |        5 |"
    "     59990 |     -49994|"
)
CONT_07501068898989 = (
    " |      |                            |          |07501068898989 |"
    "                                      |          |          |          |          |           |           |"
)


def _fixture_lines() -> list[str]:
    return FIXTURE.read_text(encoding="utf-8").splitlines()


def test_parsea_un_item_completo():
    report = parse_rows(_fixture_lines())
    row = report.rows[0]
    assert row.area == 1
    assert row.rubro == "7881 MEDIANA"
    assert row.plu == "26726760"
    assert row.barcode == "00200026726766"
    assert row.desc == "AMERICAN TOURISTER MALETA FRANKFURT."
    assert (row.existencia, row.disponible, row.precio, row.costo) == (-1, -1, 59990, -49994)


def test_ignora_encabezados_y_reglas():
    report = parse_rows(_fixture_lines())
    assert report.unknown_lines == 0
    assert len(report.rows) == 3


def test_conserva_ceros_del_codigo():
    report = parse_rows(_fixture_lines())
    assert report.rows[0].barcode.startswith("0020")


def test_asocia_codigo_alternativo_al_item_anterior():
    report = parse_rows(_fixture_lines())
    assert report.rows[0].alt_barcodes == ["07501068898989"]
    assert report.rows[1].alt_barcodes == []


def test_lee_metadatos_del_encabezado():
    report = parse_rows(_fixture_lines())
    assert report.dependencia == "10 MAIPU"
    assert report.fecha == "17/06/2026"
    assert report.hora == "10:06:36"


def test_continuacion_sin_item_previo_cuenta_como_desconocida():
    report = parse_rows([CONT_07501068898989, ITEM_26726760])
    assert report.unknown_lines == 1
    assert report.rows[0].alt_barcodes == []


def test_salta_encabezados_de_pagina_intermedios():
    report = parse_rows(_fixture_lines())
    assert len(report.rows) == 3
    assert report.unknown_lines == 0

REAL = Path(r"C:\Users\victus\Downloads\stock17-06.xlsx")


@pytest.mark.slow
@pytest.mark.skipif(not REAL.exists(), reason="archivo real no disponible")
def test_reporte_real_completo():
    report = read_report(REAL)
    assert len(report.rows) == 26527
    assert report.unknown_lines == 0
    assert len({row.area for row in report.rows}) == 69
    assert len(report.areas()) == 68  # el área 48 no tiene stock disponible
    assert sum(len(r.alt_barcodes) for r in report.rows) == 3192
    assert any(r.plu == "23842742" for r in report.rows)  # fila desbordada a B
    assert (report.dependencia, report.fecha, report.hora) == ("10 MAIPU", "17/06/2026", "10:06:36")
