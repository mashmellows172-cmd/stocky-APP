from stocky.presets import Config, Formato, load_config, save_config


def test_defaults_del_formato():
    f = Formato()
    assert (f.font_family, f.font_size, f.margin_cm, f.orientation) == ("Segoe UI", 10, 1.5, "vertical")


def test_from_dict_recorta_valores_fuera_de_rango():
    f = Formato.from_dict({"font_size": 99, "margin_cm": 9, "orientation": "diagonal"})
    assert f.font_size == 16 and f.margin_cm == 2.0 and f.orientation == "vertical"


def test_ida_y_vuelta_a_json(tmp_path):
    cfg = Config()
    cfg.presets["Sala"] = Formato(font_size=12)
    cfg.last_areas = [9, 99]
    save_config(cfg, tmp_path / "config.json")
    again = load_config(tmp_path / "config.json")
    assert again.presets["Sala"].font_size == 12 and again.last_areas == [9, 99]


def test_config_corrupta_devuelve_defaults(tmp_path):
    p = tmp_path / "config.json"
    p.write_text("{ roto", encoding="utf-8")
    cfg = load_config(p)
    assert cfg.presets == {} and cfg.last_format.font_size == 10


def test_guardado_es_atomico_sin_dejar_temporales(tmp_path):
    save_config(Config(), tmp_path / "config.json")
    assert [p.name for p in tmp_path.iterdir()] == ["config.json"]
