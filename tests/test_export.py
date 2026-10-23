import pytest

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


def test_alto_de_fila_acompana_la_letra(report_mini):
    ws = build_workbook(rows_for(report_mini, 9), 9, report_mini, Formato(font_size=16)).active
    assert ws.row_dimensions[5].height >= 19  # ALFA, primera fila de datos
    assert ws.row_dimensions[6].height >= 14  # su código alternativo


def test_sin_bordes_conserva_la_salida_actual(report_mini):
    ws = build_workbook(rows_for(report_mini, 9), 9, report_mini, Formato()).active
    assert ws["A4"].border.bottom.style == "thin"
    assert ws["A5"].border.left is None
    assert ws["A5"].border.bottom is None


@pytest.mark.parametrize(
    ("estilo", "openpyxl_style"),
    [("fino", "thin"), ("medio", "medium"), ("grueso", "thick"), ("punteado", "dotted")],
)
def test_estilos_de_celda_y_contorno(report_mini, estilo, openpyxl_style):
    fmt = Formato(cell_border=estilo, table_border=estilo)
    ws = build_workbook(rows_for(report_mini, 9), 9, report_mini, fmt).active
    assert ws["A4"].border.top.style == openpyxl_style
    assert ws["A5"].border.left.style == openpyxl_style
    assert ws["B5"].border.left.style == openpyxl_style
    assert ws.cell(ws.max_row, ws.max_column).border.bottom.style == openpyxl_style


def test_celda_y_contorno_independientes(report_mini):
    fmt = Formato(cell_border="fino", table_border="grueso")
    ws = build_workbook(rows_for(report_mini, 9), 9, report_mini, fmt).active
    assert ws["A4"].border.left.style == "thick"
    assert ws["A4"].border.top.style == "thick"
    assert ws["A4"].border.bottom.style == "thin"
    assert ws["B4"].border.left.style == "thin"
    assert ws["E4"].border.right.style == "thick"
    assert ws["A5"].border.left.style == "thick"
    assert ws["B5"].border.left.style == "thin"
    assert ws["B5"].border.right.style == "thin"


def test_codigo_alternativo_tambien_lleva_grilla(report_mini):
    fmt = Formato(cell_border="fino", table_border="grueso")
    ws = build_workbook(rows_for(report_mini, 9), 9, report_mini, fmt).active
    # ALFA es la primera fila y su código alternativo queda en la fila 6.
    assert ws["A6"].border.left.style == "thick"
    assert ws["B6"].border.right.style == "thin"
    assert ws["E6"].border.right.style == "thick"
