import tkinter as tk

import stocky.ui as ui
from stocky.pdf import PdfError
from stocky.ui import App


def test_la_ventana_muestra_el_reporte_y_avisa_si_falla_el_pdf(report_mini, tmp_path, monkeypatch):
    monkeypatch.setenv("APPDATA", str(tmp_path))
    app = App(report_mini, tmp_path / "stock17-06.xlsx", root=tk.Tk())
    try:
        assert app.selection == [9]
        assert app.current_area == 9

        monkeypatch.setattr(ui, "write_area_xlsx", lambda *args, **kwargs: tmp_path / "x.xlsx")

        def boom(*args, **kwargs):
            raise PdfError("sin Excel")

        monkeypatch.setattr(ui, "export_pdf", boom)
        app._render_worker(9, app.fmt, app._job, 1.0)
        app.root.update()
        assert "sin Excel" in app.status_render.cget("text")

        monkeypatch.setattr(ui, "export_pdf", lambda *args, **kwargs: (_ for _ in ()).throw(RuntimeError("raro")))
        app._render_worker(9, app.fmt, app._job, 1.0)
        app.root.update()
        assert "raro" in app.status_render.cget("text")

        monkeypatch.setattr(ui, "write_area_xlsx", lambda *args, **kwargs: tmp_path / "y.xlsx")
        monkeypatch.setattr(ui, "print_xlsx", lambda xlsx: None)
        app._print_worker(9, app.fmt)
        app.root.update()
        assert "enviada a la impresora" in app.status_render.cget("text")

        def sin_impresora(xlsx):
            raise PdfError("sin impresora")

        monkeypatch.setattr(ui, "print_xlsx", sin_impresora)
        app._print_worker(9, app.fmt)
        app.root.update()
        assert "sin impresora" in app.status_render.cget("text")
    finally:
        app.root.destroy()
