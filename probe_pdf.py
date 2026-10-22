import sys
sys.path.insert(0, ".")
from pathlib import Path
import tempfile
from conftest import _row
from stocky.models import StockReport
from stocky.export import build_workbook, rows_for
from stocky.presets import Formato
from stocky import pdf as pdfmod

rows=[_row(9,'ALFA',1,'00001234','00200026726766',59990,['07501068898989'])]
r=StockReport(fecha='17/06/2026',hora='10:06:36',dependencia='10 MAIPU',rows=rows)
d=Path(tempfile.mkdtemp()); p=d/'m.xlsx'
build_workbook(rows_for(r,9),9,r,Formato()).save(p)

import win32com.client, pywintypes
app=win32com.client.DispatchEx('Excel.Application')
print('dispatch ok', flush=True)
app.Visible=False; app.DisplayAlerts=False
print('props ok', flush=True)
wb=app.Workbooks.Open(str(p.resolve()), ReadOnly=True)
print('open ok', flush=True)
wb.ExportAsFixedFormat(0,str((d/'m.pdf').resolve()))
print('export ok', flush=True)
wb.Close(False); print('close ok', flush=True)
app.Quit(); print('quit ok', flush=True)
