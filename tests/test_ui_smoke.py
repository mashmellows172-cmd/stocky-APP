import tkinter as tk

from stocky.ui import App


def test_la_ventana_se_construye_con_un_reporte(report_mini, tmp_path, monkeypatch):
    monkeypatch.setenv("APPDATA", str(tmp_path))
    app = App(report_mini, tmp_path / "stock17-06.xlsx", root=tk.Tk())
    try:
        assert app.selection == [9]
        assert app.current_area == 9
    finally:
        app.root.destroy()
