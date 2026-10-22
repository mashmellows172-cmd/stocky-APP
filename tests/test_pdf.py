import pytest

from stocky.pdf import PdfError, excel_available, export_pdf


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
