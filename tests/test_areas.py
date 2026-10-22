from stocky.areas import AREA_NAMES, area_name, file_stem, header_label


def test_mapa_completo_del_letrero():
    assert area_name(9) == "TELEFONIA"
    assert area_name(99) == "TELEVISION"
    assert area_name(80) == "HORNO"
    assert len(AREA_NAMES) == 18


def test_area_fuera_del_mapa():
    assert area_name(37) is None
    assert file_stem(37) == "37"
    assert header_label(37) == "ÁREA 37"


def test_nombre_de_archivo():
    assert file_stem(9) == "09 - TELEFONIA"
    assert header_label(9) == "ÁREA 09 · TELEFONIA"
