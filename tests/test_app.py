import stocky.app as app
from stocky.app import main
from stocky.parser import ParseError


def test_sin_argumentos_explica_el_uso(capsys):
    assert main([]) == 2
    assert "arrastra" in capsys.readouterr().err.lower()


def test_archivo_inexistente_devuelve_error(capsys, tmp_path):
    assert main([str(tmp_path / "no.xlsx")]) == 2
    assert "no existe" in capsys.readouterr().err.lower()


def test_reporte_vacio_muestra_el_mensaje(monkeypatch, capsys, tmp_path):
    monkeypatch.setattr(app, "read_report", lambda p: (_ for _ in ()).throw(ParseError("El reporte no trae ítems")))
    assert main([str(tmp_path)]) == 1
    assert "no trae ítems" in capsys.readouterr().err


def test_archivo_que_no_es_excel_avisa(capsys, tmp_path):
    bad = tmp_path / "malo.xlsx"
    bad.write_text("no soy un excel", encoding="utf-8")
    assert main([str(bad)]) == 1
    assert "no se pudo leer" in capsys.readouterr().err.lower()
