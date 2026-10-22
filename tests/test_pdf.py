import pytest

from stocky.pdf import PdfError, excel_available, export_pdf, print_xlsx


@pytest.mark.skipif(not excel_available(), reason="sin Excel")
def test_exporta_pdf_de_una_pagina(xlsx_mini, tmp_path):
    out = export_pdf(xlsx_mini, tmp_path / "a.pdf")
    assert out.exists() and out.stat().st_size > 1000
    import pypdfium2 as pdfium

    document = pdfium.PdfDocument(str(out))
    try:
        assert len(document) == 1
    finally:
        document.close()


def test_error_si_el_xlsx_no_existe(tmp_path):
    with pytest.raises(PdfError, match="no existe"):
        export_pdf(tmp_path / "nada.xlsx", tmp_path / "a.pdf")


def test_imprimir_archivo_inexistente(tmp_path):
    with pytest.raises(PdfError, match="no existe"):
        print_xlsx(tmp_path / "nada.xlsx")


def test_imprimir_usa_excel_y_cierra(tmp_path, monkeypatch):
    import stocky.pdf as pdfmod

    xlsx = tmp_path / "a.xlsx"
    xlsx.write_bytes(b"x")
    calls: list[str] = []

    class FakeWorkbook:
        def PrintOut(self, *args, **kwargs):
            calls.append("print")

        def Close(self, *args):
            calls.append("close")

    class FakeBooks:
        def Open(self, *args, **kwargs):
            calls.append("open")
            return FakeWorkbook()

    class FakeApp:
        Visible = True
        DisplayAlerts = True
        Workbooks = FakeBooks()

        def Quit(self):
            calls.append("quit")

    monkeypatch.setattr(pdfmod, "_excel", lambda: FakeApp())
    print_xlsx(xlsx)
    assert calls == ["open", "print", "close", "quit"]
