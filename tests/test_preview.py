import io

import stocky.preview as preview
from stocky.preview import clear_cache, render_pages


def test_renderiza_una_pagina_como_png(pdf_mini):
    pages = render_pages(pdf_mini)
    assert len(pages) == 1
    assert pages[0][:8] == b"\x89PNG\r\n\x1a\n"


def test_cache_no_relee_el_mismo_pdf(pdf_mini, monkeypatch):
    clear_cache()
    primera = render_pages(pdf_mini)

    def explode(path):
        raise AssertionError("releyó")

    monkeypatch.setattr(preview, "_open_pdf", explode)
    assert render_pages(pdf_mini) == primera


def test_escala_mayor_da_mas_pixeles(pdf_mini):
    from PIL import Image

    chica = Image.open(io.BytesIO(render_pages(pdf_mini, scale=1.0)[0])).size
    grande = Image.open(io.BytesIO(render_pages(pdf_mini, scale=2.0)[0])).size
    assert grande[0] > chica[0]
