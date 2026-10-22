from stocky.export import build_workbook, rows_for, write_area_xlsx
from stocky.presets import Formato


def test_filtra_y_ordena(report_mini):
    rows = rows_for(report_mini, area=9)
    assert [r.desc for r in rows] == ["ALFA", "BETA", "GAMMA"]
    assert all(r.disponible >= 1 for r in rows)


def test_codigo_y_plu_quedan_como_texto(report_mini):
    ws = build_workbook(rows_for(report_mini, 9), 9, report_mini, Formato()).active
    codigos = [c.value for c in ws["B"][4:]]
    assert "00200026726766" in codigos


def test_configuracion_de_pagina_carta_y_una_hoja_de_ancho(report_mini):
    ws = build_workbook(
        rows_for(report_mini, 9), 9, report_mini, Formato(font_size=16, wide_desc=True)
    ).active
    assert ws.page_setup.paperSize == int(ws.PAPERSIZE_LETTER)
    assert ws.page_setup.fitToWidth == 1
    assert ws.sheet_properties.pageSetUpPr.fitToPage is True
    assert ws.print_title_rows == "$1:$4"


def test_pie_con_ajuste_y_totales(report_mini):
    ws = build_workbook(rows_for(report_mini, 9), 9, report_mini, Formato()).active
    assert "Página &P de &N" in ws.oddFooter.right.text
    assert "3 ítems" in ws.oddFooter.left.text


def test_area_sin_stock_no_escribe_archivo(report_mini, tmp_path):
    assert write_area_xlsx(report_mini, area=8, fmt=Formato(), out_dir=tmp_path) is None
