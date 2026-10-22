"""Entrada de Stocky: arrastrar el reporte sobre el ícono o pasarlo como argumento."""

from __future__ import annotations

import sys
from pathlib import Path

if __package__ in (None, ""):  # ejecutado como script por el acceso directo
    sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from stocky import ui
from stocky.parser import ParseError, read_report

USO = (
    "Arrastra el archivo del reporte (.xlsx) sobre el ícono de Stocky, "
    "o pásalo como argumento: python stocky/app.py <reporte.xlsx>"
)


def _show_error(message: str) -> None:
    try:
        from tkinter import Tk, messagebox

        root = Tk()
        root.withdraw()
        messagebox.showerror("Stocky", message)
        root.destroy()
    except Exception:  # pragma: no cover - sin entorno gráfico
        pass


def _error(message: str) -> None:
    if sys.stderr is not None:
        print(message, file=sys.stderr)
    else:  # pythonw: no hay consola donde leerlo
        _show_error(message)


def main(argv: list[str]) -> int:
    if not argv:
        _error(USO)
        return 2

    path = Path(argv[0])
    if not path.exists():
        _error(f"El archivo no existe: {path}")
        return 2

    try:
        report = read_report(path)
    except ParseError as error:
        _error(str(error))
        return 1
    except (OSError, ValueError) as error:
        _error(f"No se pudo leer el reporte: {error}")
        return 1

    ui.run(report, path)
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
