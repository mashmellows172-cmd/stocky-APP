import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).parent))

from stocky.models import StockReport, StockRow  # noqa: E402


def _row(area: int, desc: str, disponible: int, plu: str, barcode: str, precio: int = 1000, alt=None) -> StockRow:
    return StockRow(
        area=area,
        rubro="7881 MEDIANA",
        plu=plu,
        barcode=barcode,
        desc=desc,
        un_emp=disponible,
        un_despa=disponible,
        existencia=disponible,
        disponible=disponible,
        precio=precio,
        costo=0,
        alt_barcodes=list(alt or []),
    )


@pytest.fixture
def report_mini() -> StockReport:
    rows = [
        _row(9, "GAMMA", 2, "00003456", "00200000003456", 1500),
        _row(9, "ALFA", 1, "00001234", "00200026726766", 59990, ["07501068898989"]),
        _row(9, "BETA", 1, "00002345", "00200026726799", 37990),
        _row(9, "ZETA", 0, "00007890", "00200000007890"),
        _row(8, "DELTA", 0, "00009999", "00200000009999"),
    ]
    return StockReport(dependencia="10 MAIPU", fecha="17/06/2026", hora="10:06:36", rows=rows)
