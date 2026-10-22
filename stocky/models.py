from dataclasses import dataclass, field
from pathlib import Path


@dataclass(slots=True)
class StockRow:
    area: int
    rubro: str
    plu: str
    barcode: str
    desc: str
    un_emp: int
    un_despa: int
    existencia: int
    disponible: int
    precio: int
    costo: int
    alt_barcodes: list[str] = field(default_factory=list)


@dataclass(slots=True)
class StockReport:
    dependencia: str = ""
    fecha: str = ""
    hora: str = ""
    rows: list[StockRow] = field(default_factory=list)
    unknown_lines: int = 0
    conflicts: int = 0
    source: Path | None = None

    def areas(self) -> dict[int, int]:
        """Área -> cantidad de ítems con stock disponible."""
        counts: dict[int, int] = {}
        for row in self.rows:
            if row.disponible >= 1:
                counts[row.area] = counts.get(row.area, 0) + 1
        return dict(sorted(counts.items()))
