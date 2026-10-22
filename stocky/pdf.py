"""Abrir Excel, exportar a PDF y enviar a la impresora."""

from __future__ import annotations

from pathlib import Path


class PdfError(Exception):
    """Excel no pudo abrir, convertir o imprimir el archivo."""


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


def _run_in_excel(xlsx: Path, action, error_message: str):
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
        workbook = app.Workbooks.Open(str(Path(xlsx).resolve()), ReadOnly=True)
        return action(app, workbook)
    except pywintypes.com_error as error:
        raise PdfError(f"{error_message}: {error}") from error
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


def export_pdf(xlsx: Path, pdf: Path) -> Path:
    xlsx = Path(xlsx)
    pdf = Path(pdf)
    if not xlsx.exists():
        raise PdfError(f"El archivo no existe: {xlsx}")
    pdf.parent.mkdir(parents=True, exist_ok=True)
    target = str(pdf.resolve())
    _run_in_excel(
        xlsx,
        lambda app, workbook: workbook.ExportAsFixedFormat(0, target),
        "Excel no pudo exportar el PDF",
    )
    return pdf


def print_xlsx(xlsx: Path) -> None:
    """Imprime el libro con Excel (el verbo «print» del shell no existe para el visor de PDF)."""
    xlsx = Path(xlsx)
    if not xlsx.exists():
        raise PdfError(f"El archivo no existe: {xlsx}")
    _run_in_excel(xlsx, lambda app, workbook: workbook.PrintOut(), "No se pudo imprimir")
