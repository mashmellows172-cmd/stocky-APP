"""Crea el ícono y el acceso directo Stocky.lnk para arrastrar el reporte."""

from __future__ import annotations

import sys
from pathlib import Path

PROJECT = Path(__file__).resolve().parents[1]
ASSETS = PROJECT / "assets"
ICON = ASSETS / "stocky.ico"
LINK = PROJECT / "Stocky.lnk"


def crear_icono(destino: Path = ICON) -> Path:
    from PIL import Image, ImageDraw

    size = 256
    image = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    draw = ImageDraw.Draw(image)
    draw.rounded_rectangle((8, 8, size - 8, size - 8), radius=44, fill="#059669")

    bars = [(58, 150, 92, 200), (110, 108, 144, 200), (162, 66, 196, 200)]
    for left, top, right, bottom in bars:
        draw.rounded_rectangle((left, top, right, bottom), radius=8, fill="#FFFFFF")
    draw.rectangle((52, 204, 196, 212), fill="#FFFFFF")

    destino.parent.mkdir(parents=True, exist_ok=True)
    image.save(destino, sizes=[(16, 16), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)])
    return destino


def crear_acceso(destino: Path = LINK, icono: Path = ICON) -> Path:
    import win32com.client

    icono = crear_icono(icono)
    pythonw = Path(sys.executable).with_name("pythonw.exe")
    target = pythonw if pythonw.exists() else Path(sys.executable)

    shell = win32com.client.Dispatch("WScript.Shell")
    link = shell.CreateShortCut(str(destino))
    link.Targetpath = str(target)
    link.Arguments = f'"{PROJECT / "stocky" / "app.py"}"'
    link.IconLocation = str(icono)
    link.WorkingDirectory = str(PROJECT)
    link.Description = "Stocky - tablas de stock por área"
    link.save()
    return destino


def main() -> int:
    icono = crear_icono()
    acceso = crear_acceso(icono=icono)
    print(f"Ícono: {icono}")
    print(f"Acceso directo: {acceso}")
    print("Arrastra el reporte .xlsx sobre el acceso directo para abrir el panel.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
