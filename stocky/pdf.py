"""Exportación a PDF con Excel y envío a la impresora."""

from __future__ import annotations

import os
from pathlib import Path


class PdfError(Exception):
    """Excel no pudo abrir o convertir el archivo."""


def _excel():
    import win32com.client

    return win32com.client.DispatchEx("Excel.Application")


def excel_available() -> bool:
    """True si Excel está instalado (evita abrirlo sólo para preguntar)."""
    try:
        import winreg
    except ImportError:
        return False
    try:
        winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, "Excel.Application").Close()
    except OSError:
        return False
    return True


def export_pdf(xlsx: Path, pdf: Path) -> Path:
    xlsx = Path(xlsx)
    pdf = Path(pdf)
    if not xlsx.exists():
        raise PdfError(f"El archivo no existe: {xlsx}")
    pdf.parent.mkdir(parents=True, exist_ok=True)

    try:
        import pythoncom
        import pywintypes
    except ImportError as error:  # pragma: no cover - sólo sin pywin32
        raise PdfError(f"Falta pywin32: {error}") from error

    pythoncom.CoInitialize()
    app = None
    workbook = None
    try:
        app = _excel()
        app.Visible = False
        app.DisplayAlerts = False
        workbook = app.Workbooks.Open(str(xlsx.resolve()), ReadOnly=True)
        workbook.ExportAsFixedFormat(0, str(pdf.resolve()))  # 0 = xlTypePDF
    except pywintypes.com_error as error:
        raise PdfError(f"Excel no pudo exportar el PDF: {error}") from error
    finally:
        if workbook is not None:
            try:
                workbook.Close(False)
            except Exception:
                pass
        if app is not None:
            try:
                app.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()
    return pdf
    return pdf


def print_pdf(pdf: Path) -> None:
    pdf = Path(pdf)
    if not pdf.exists():
        raise PdfError(f"El archivo no existe: {pdf}")
    try:
        os.startfile(str(pdf), "print")
    except OSError as error:
        raise PdfError(f"No se pudo enviar a la impresora: {error}") from error
