"""Render de las páginas del PDF para la vista previa."""

from __future__ import annotations

import io
from pathlib import Path

import pypdfium2 as pdfium

MAX_CACHE = 20

_cache: dict[tuple[str, int, float], list[bytes]] = {}


def clear_cache() -> None:
    _cache.clear()


def _open_pdf(path: Path | str):
    return pdfium.PdfDocument(str(path))


def render_pages(pdf: Path | str, scale: float = 1.5) -> list[bytes]:
    """PNG por página del PDF, en orden."""
    path = Path(pdf)
    key = (str(path.resolve()), path.stat().st_mtime_ns, scale)
    cached = _cache.get(key)
    if cached is not None:
        return cached

    document = _open_pdf(path)
    try:
        pages: list[bytes] = []
        for page in document:
            bitmap = page.render(scale=scale)
            image = bitmap.to_pil().convert("RGB")
            buffer = io.BytesIO()
            image.save(buffer, format="PNG")
            pages.append(buffer.getvalue())
    finally:
        document.close()

    _cache[key] = pages
    while len(_cache) > MAX_CACHE:
        _cache.pop(next(iter(_cache)), None)
    return pages
