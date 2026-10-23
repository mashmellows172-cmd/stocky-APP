"""Formato de impresión y presets persistentes."""

from __future__ import annotations

import json
import os
from dataclasses import asdict, dataclass, field, fields
from pathlib import Path

DEFAULT_CONFIG_PATH = Path(os.environ.get("APPDATA", Path.home())) / "Stocky" / "config.json"

MARGINS = (0.8, 1.0, 1.5, 2.0)
DENSITIES = ("compacta", "normal", "amplia")
ORIENTATIONS = ("vertical", "horizontal")
HEADER_FILLS = ("ninguna", "gris", "verde")
BORDER_STYLES = ("ninguno", "fino", "medio", "grueso", "punteado")


@dataclass(slots=True)
class Formato:
    font_family: str = "Segoe UI"
    font_size: int = 10
    title_size: int = 14
    density: str = "normal"
    zebra: bool = False
    header_fill: str = "gris"
    margin_cm: float = 1.5
    orientation: str = "vertical"
    wide_desc: bool = False
    show_rubro: bool = False
    thousands: bool = True
    extra_header: str = ""
    min_stock: int = 1
    cell_border: str = "ninguno"
    table_border: str = "ninguno"

    def to_dict(self) -> dict:
        return asdict(self)

    @classmethod
    def from_dict(cls, data: dict) -> "Formato":
        if not isinstance(data, dict):
            return cls()
        base = cls()
        out = {}
        for item in fields(cls):
            raw = data.get(item.name, getattr(base, item.name))
            out[item.name] = _sanitize(item.name, raw, getattr(base, item.name))
        return cls(**out)


def _sanitize(name: str, value, default):
    if name in ("font_size", "title_size"):
        return _clamp_int(value, 8, 16 if name == "font_size" else 24, default)
    if name == "min_stock":
        return _clamp_int(value, 0, 999, default)
    if name == "margin_cm":
        try:
            value = float(value)
        except (TypeError, ValueError):
            return default
        if value in MARGINS:
            return value
        value = max(MARGINS[0], min(MARGINS[-1], value))
        return min(MARGINS, key=lambda margin: abs(margin - value))
    if name in ("zebra", "wide_desc", "show_rubro", "thousands"):
        return value if isinstance(value, bool) else default
    if name == "density":
        return value if value in DENSITIES else default
    if name == "orientation":
        return value if value in ORIENTATIONS else default
    if name == "header_fill":
        return value if value in HEADER_FILLS else default
    if name in ("cell_border", "table_border"):
        return value if value in BORDER_STYLES else default
    if name == "font_family":
        return value.strip() if isinstance(value, str) and value.strip() else default
    if name == "extra_header":
        return value if isinstance(value, str) else default
    return value


def _clamp_int(value, low: int, high: int, default: int) -> int:
    try:
        number = int(value)
    except (TypeError, ValueError):
        return default
    return max(low, min(high, number))


@dataclass(slots=True)
class Config:
    last_format: Formato = field(default_factory=Formato)
    presets: dict[str, Formato] = field(default_factory=dict)
    last_areas: list[int] = field(default_factory=list)
    last_folder: str = ""

    def to_dict(self) -> dict:
        return {
            "last_format": self.last_format.to_dict(),
            "presets": {name: fmt.to_dict() for name, fmt in self.presets.items()},
            "last_areas": list(self.last_areas),
            "last_folder": self.last_folder,
        }

    @classmethod
    def from_dict(cls, data: dict) -> "Config":
        if not isinstance(data, dict):
            return cls()
        presets_raw = data.get("presets")
        presets = {}
        if isinstance(presets_raw, dict):
            for name, fmt in presets_raw.items():
                if isinstance(name, str) and isinstance(fmt, dict):
                    presets[name] = Formato.from_dict(fmt)
        areas = data.get("last_areas")
        last_areas = [a for a in areas if isinstance(a, int)] if isinstance(areas, list) else []
        folder = data.get("last_folder")
        return cls(
            last_format=Formato.from_dict(data.get("last_format") or {}),
            presets=presets,
            last_areas=last_areas,
            last_folder=folder if isinstance(folder, str) else "",
        )


def load_config(path: Path = DEFAULT_CONFIG_PATH) -> Config:
    try:
        data = json.loads(Path(path).read_text(encoding="utf-8"))
    except (FileNotFoundError, json.JSONDecodeError, TypeError, OSError):
        return Config()
    try:
        return Config.from_dict(data)
    except (TypeError, ValueError):
        return Config()


def save_config(cfg: Config, path: Path = DEFAULT_CONFIG_PATH) -> None:
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_suffix(path.suffix + ".tmp")
    tmp.write_text(json.dumps(cfg.to_dict(), ensure_ascii=False, indent=2), encoding="utf-8")
    tmp.replace(path)
