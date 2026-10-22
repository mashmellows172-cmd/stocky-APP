"""Mapa de áreas del letrero de la tienda."""

from __future__ import annotations

AREA_NAMES: dict[int, str] = {
    8: "AUDIO",
    9: "TELEFONIA",
    10: "ELECTRODOMESTICOS",
    29: "ELECTRONICA MENOR",
    46: "COMPUTACION",
    73: "ACCESORIOS COMPUTACION",
    74: "ENTRETENCION",
    77: "CUIDADO PERSONAL",
    78: "PYMES",
    79: "REFRIGERACION",
    80: "HORNO",
    81: "LAVADO Y SECADO",
    83: "ASPIRADO Y LIMPIEZA",
    84: "COCINA",
    85: "OTROS LINEA BLANCA",
    86: "PERIFERICOS",
    95: "TEMPORADA",
    99: "TELEVISION",
}


def area_name(area: int) -> str | None:
    return AREA_NAMES.get(area)


def file_stem(area: int) -> str:
    """Nombre de archivo sin extensión: ``09 - TELEFONIA`` o ``37``."""
    name = area_name(area)
    return f"{area:02d} - {name}" if name else str(area)


def header_label(area: int) -> str:
    """Encabezado impreso: ``ÁREA 09 · TELEFONIA`` o ``ÁREA 37``."""
    name = area_name(area)
    return f"ÁREA {area:02d} · {name}" if name else f"ÁREA {area}"
