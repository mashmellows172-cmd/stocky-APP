"""Construye Stocky.exe como un único archivo portable para Windows 64 bits."""

from __future__ import annotations

import subprocess
import sys
from pathlib import Path

PROJECT = Path(__file__).resolve().parents[1]
ENTRY = PROJECT / "stocky" / "app.py"
ICON = PROJECT / "assets" / "stocky.ico"
DIST = PROJECT / "dist"
BUILD = PROJECT / "build"
EXE = DIST / "Stocky.exe"

LEEME = """Stocky - versión portable

1. Copiá Stocky.exe a la PC donde lo vas a usar.
2. Arrastrá el reporte .xlsx sobre Stocky.exe para abrirlo.
   También podés ejecutarlo desde una consola pasándole la ruta del reporte.
3. No hace falta instalar Python: el ejecutable ya lo incluye.

Para generar el PDF o imprimir, la PC necesita Microsoft Excel instalado.
Guardar el .xlsx funciona sin Excel.

Si Windows muestra un aviso de SmartScreen por ser un programa nuevo sin
firma, elegí "Más información" y después "Ejecutar de todas formas".
"""


def construir_exe() -> Path:
    if not ENTRY.exists():
        raise FileNotFoundError(f"No existe el punto de entrada: {ENTRY}")
    if not ICON.exists():
        raise FileNotFoundError(f"No existe el icono: {ICON}")

    comando = [
        sys.executable,
        "-m",
        "PyInstaller",
        "--noconfirm",
        "--clean",
        "--onefile",
        "--windowed",
        "--name",
        "Stocky",
        "--icon",
        str(ICON),
        "--distpath",
        str(DIST),
        "--workpath",
        str(BUILD),
        "--specpath",
        str(BUILD),
        "--hidden-import",
        "pythoncom",
        "--hidden-import",
        "pywintypes",
        "--hidden-import",
        "win32timezone",
        "--collect-all",
        "pypdfium2",
        str(ENTRY),
    ]
    print("Construyendo Stocky.exe...")
    subprocess.run(comando, cwd=PROJECT, check=True)

    if not EXE.exists():
        raise FileNotFoundError(f"PyInstaller no generó el ejecutable: {EXE}")

    (DIST / "LEEME Stocky.txt").write_text(LEEME, encoding="utf-8")
    print(f"Ejecutable: {EXE} ({EXE.stat().st_size:,} bytes)")
    return EXE


def main() -> int:
    construir_exe()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
