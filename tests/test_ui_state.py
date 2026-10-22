from stocky.ui_state import DEFAULT_AREAS, initial_selection, resolve_font, status_text


def test_seleccion_inicial_usa_las_seis_de_siempre(report_mini):
    assert DEFAULT_AREAS == [8, 9, 29, 46, 84, 99]
    assert initial_selection(sorted(report_mini.areas()), []) == [9]  # solo 9 tiene stock en el mini


def test_seleccion_inicial_respeta_la_ultima(report_mini):
    assert initial_selection([8, 9, 29], [29]) == [29]


def test_seleccion_ignora_areas_que_no_existen(report_mini):
    assert initial_selection([9], [9, 99]) == [9]


def test_seleccion_sin_coincidencias_usa_todas(report_mini):
    assert initial_selection([37, 44], [99]) == [37, 44]


def test_status_resume_anomalias(report_mini):
    report_mini.unknown_lines = 2
    report_mini.conflicts = 1
    texto = status_text(report_mini, [9])
    assert "2 líneas no reconocidas" in texto and "1 duplicado conflictivo" in texto
    assert "5 ítems" in texto and "1 área" in texto


def test_fuente_inexistente_cae_en_segoe_ui():
    assert resolve_font("NoExiste", ["Segoe UI", "Arial"]) == "Segoe UI"
    assert resolve_font("Arial", ["Segoe UI", "Arial"]) == "Arial"
