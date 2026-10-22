from pathlib import Path

from stocky.parser import parse_rows

FIXTURE = Path(__file__).parent / "fixtures" / "reporte_recorte.txt"


def test_parsea_un_item_completo():
    report = parse_rows(FIXTURE.read_text(encoding="utf-8").splitlines())
    row = report.rows[0]
    assert row.area == 1
    assert row.rubro == "7881 MEDIANA"
    assert row.plu == "26726760"
    assert row.barcode == "00200026726766"
    assert row.desc == "AMERICAN TOURISTER MALETA FRANKFURT."
    assert (row.existencia, row.disponible, row.precio, row.costo) == (-1, -1, 59990, -49994)


def test_ignora_encabezados_y_reglas():
    report = parse_rows(FIXTURE.read_text(encoding="utf-8").splitlines())
    assert report.unknown_lines == 0
    assert len(report.rows) == 2


def test_conserva_ceros_del_codigo():
    report = parse_rows(FIXTURE.read_text(encoding="utf-8").splitlines())
    assert report.rows[0].barcode.startswith("0020")
