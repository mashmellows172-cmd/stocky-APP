# Stocky — panel de extracción y formato de stock · Plan de implementación

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Convertir el reporte diario de stock del SAI en tablas por área listas para imprimir, con panel de formato y vista previa real.

**Architecture:** Un parser puro transforma el texto del xlsx en `StockReport`; `export.py` arma el xlsx por área con openpyxl; Excel (COM) lo convierte a PDF y ese PDF alimenta la vista previa (pypdfium2) y la impresión. La UI Tkinter solo orquesta; toda la lógica es testeable sin ventanas.

**Tech Stack:** Python 3.14, Tkinter 8.6, openpyxl, pypdfium2 + Pillow, pywin32 (COM de Excel), pytest.

**Spec:** `docs/superpowers/specs/2026-10-22-stocky-panel-design.md`

## Global Constraints

- Windows 10/11. Python 3.14 (`py` o `python`). Sin red en tiempo de ejecución.
- Papel **carta (Letter)**; vertical por defecto; `fitToWidth = 1` siempre: nunca dos páginas de ancho.
- Tabla: `PLU · CÓDIGO · DESCRIPCIÓN · DISPONIBLE · PRECIO` (rubro opcional, apagado).
- Solo `disponible >= 1`; orden ascendente por descripción (case/acentos ignorados), desempate PLU.
- PLU y código de barra son **texto**: los ceros a la izquierda no se pierden.
- Nunca se modifica ni se mueve el archivo original.
- Salida: `<carpeta del original>\Stock YYYY-MM-DD\NN - NOMBRE.xlsx` y `.pdf`.
- Dependencias de ejecución: `openpyxl`, `pypdfium2`, `Pillow`, `pywin32`. Desarrollo: `pytest`.
- Config en `%APPDATA%\Stocky\config.json`; nunca dentro del proyecto.
- Fuente por defecto Segoe UI 10; márgenes 1,5 cm; precios con separador de miles.

## Review Focus

1. **Fila desbordada a la celda B** (fila 37.067 real): el ítem debe recuperarse completo, no perderse ni contarse como línea desconocida.
2. **Reporte sin ítems** (`stock19-10.xlsx`, 44.032 filas de encabezado): error claro, sin panel y sin carpetas vacías.
3. **PLU repetido**: con otro código de barra se fusiona como alternativo; con stock/precio distinto se conserva el primero y se informa el conflicto.
4. **Área sin stock o fuera del mapa**: no se genera archivo para áreas vacías; las desconocidas usan el número como nombre (`37 - .xlsx` → `37.xlsx`).
5. **Config corrupta o fuente inexistente**: la app arranca igual con los valores por defecto.

---

### Task 1: Repo, dependencias y parser de una línea

**Files:**
- Create: `.gitignore`, `requirements.txt`, `stocky/__init__.py`, `stocky/models.py`, `stocky/parser.py`
- Create: `tests/fixtures/reporte_recorte.txt`, `tests/test_parser.py`

**Interfaces:**
- Produces:
  - `stocky.models.StockRow` (dataclass, slots): `area:int, rubro:str, plu:str, barcode:str, desc:str, un_emp:int, un_despa:int, existencia:int, disponible:int, precio:int, costo:int, alt_barcodes:list[str]`
  - `stocky.models.StockReport` (dataclass, slots): `dependencia:str, fecha:str, hora:str, rows:list[StockRow], unknown_lines:int, conflicts:int, source:pathlib.Path|None`; método `areas() -> dict[int,int]` (área → ítems con stock).
  - `stocky.parser.ParseError(Exception)`, `read_report(path: Path) -> StockReport`, `parse_rows(lines: Iterable[str]) -> StockReport`.

- [ ] **Step 1: Inicializar repo y dependencias**

```bash
git init
python -m pip install openpyxl pypdfium2 Pillow pywin32 pytest
python -m pip freeze | Select-String -Pattern "openpyxl|pypdfium2|pillow|pywin32" > requirements.txt
```

`.gitignore`: `__pycache__/`, `.pytest_cache/`, `.superpowers/`, `*.pyc`, `dist/`, `build/`.

- [ ] **Step 2: Crear el fixture con el recorte real**

`tests/fixtures/reporte_recorte.txt` (contenido literal, sin los espacios de relleno del final de cada línea; en el archivo real la línea `rea :` trae una vocal rota, aquí no importa porque no lleva `|`):

```text
  FECHA : 17/06/2026                                   L I S T A D O   DE   E X I S T E N C I A S                                      RInLisExiP
  HORA  : 10:06:36                                                                                                                     PAGINA : 1
DEPENDENCIA   :   10 MAIPU
Area          :    0 AREA: TODOS LOS ITEMS.
FECHA EMISION : 17/06/2026
  ---------------------------------------------------------------------------------------------------------------------
 | AREA |     NIVEL DE AGREGACION    |   PLU    |CODIGO DE BARRA|          DESCRIPCION PRODUCTO        |            E X I S T E N C I A S       |||| PRECIO DE |COSTO      |
 |      |                            |          |               |                                      |----------------------------------------||||   VENTA   |PROMEDIO   |
 |      |                            |          |               |                                      | UN. EMP  | UN. DESPA|EXISTENCIA|DISPONIBLE|           |           |
  ---------------------------------------------------------------------------------------------------------------------
 |  1   |7881 MEDIANA                |26726760  |00200026726766 | AMERICAN TOURISTER MALETA FRANKFURT. |(UNI)   -1|(UNI)   -1|       -1 |       -1 |     59990 |     -49994|
 |      |                            |          |07501068898989 |                                      |          |          |          |          |           |           |
 |  1   |3888 BOLSO CONCESION        |28098820  |00200028098823 |CONCESION XTREM NAVY BLUE RIDER3XT SP |(UNI)    1|(UNI)    1|        1 |        1 |     37990 |      20296|
```

- [ ] **Step 3: Escribir las pruebas que fallan**

`tests/test_parser.py`:

```python
from pathlib import Path
from stocky.parser import parse_rows

FIXTURE = Path(__file__).parent / "fixtures" / "reporte_recorte.txt"

def test_parsea_un_item_completo():
    report = parse_rows(FIXTURE.read_text(encoding="utf-8").splitlines())
    row = report.rows[0]
    assert row.area == 1
    assert row.rubro == "7881 MEDIANA"
    assert row.plu == "26726760"
    assert row.barcode == "00200026726766"
    assert row.desc == "AMERICAN TOURISTER MALETA FRANKFURT."
    assert (row.existencia, row.disponible, row.precio, row.costo) == (-1, -1, 59990, -49994)

def test_ignora_encabezados_y_reglas():
    report = parse_rows(FIXTURE.read_text(encoding="utf-8").splitlines())
    assert report.unknown_lines == 0
    assert len(report.rows) == 2

def test_conserva_ceros_del_codigo():
    report = parse_rows(FIXTURE.read_text(encoding="utf-8").splitlines())
    assert report.rows[0].barcode.startswith("0020")
```

- [ ] **Step 4: Correr y ver que fallan**

Run: `python -m pytest tests/test_parser.py -v`
Expected: FAIL con `ModuleNotFoundError: No module named 'stocky'` (o `ImportError: parse_rows`).

- [ ] **Step 5: Implementar `models.py` y el parser mínimo**

`parse_rows` parte cada línea por `|` (13 separadores), hace `strip()` de los campos, considera **ítem** cuando el campo 1 es dígito y el campo 3 es dígito, y **continuación** cuando el campo 4 es dígito y los campos 1/3 están vacíos. En esta tarea la continuación se reconoce y se descarta sin sumar a `unknown_lines`; en Task 2 pasa a adjuntarse al ítem anterior. Los enteros de cantidades salen del último token del campo (`"(UNI)   -1"` → `-1`). Toda línea con `|` que no sea ítem ni continuación suma a `unknown_lines`; las líneas sin `|` se ignoran en silencio. Si no hay ningún ítem: `raise ParseError("El reporte no trae ítems")` (se prueba en Task 3).

`read_report(path)` abre el xlsx con `openpyxl.load_workbook(path, data_only=True, read_only=True)`, concatena **todas las celdas no vacías** de cada fila (`"".join(str(c) for c in row if c is not None)`) y llama a `parse_rows`. Acepta cualquier hoja; usa la activa.

- [ ] **Step 6: Correr y ver que pasan**

Run: `python -m pytest tests/test_parser.py -v`
Expected: 3 passed.

- [ ] **Step 7: Commit**

```bash
git add .
git commit -m "feat: parser base del listado de existencias"
```

---

### Task 2: Códigos alternativos, metadatos y páginas repetidas

**Files:**
- Modify: `stocky/parser.py`
- Test: `tests/test_parser.py`

**Interfaces:**
- Consumes: `StockRow`, `StockReport` de Task 1.
- Produces: `StockRow.alt_barcodes` poblado; `StockReport.dependencia/fecha/hora` con `"10 MAIPU"`, `"17/06/2026"`, `"10:06:36"`.

- [ ] **Step 1: Escribir las pruebas que fallan**

Constantes del módulo de pruebas, con el formato exacto del fixture: `ITEM_26726760`, `ITEM_26726760_CON_CODIGO_2` (mismo ítem con `00200026726799`), `ITEM_26726760_STOCK_5` (disponible 5), `CONT_07501068898989`.

```python
def test_asocia_codigo_alternativo_al_item_anterior():
    report = parse_rows(FIXTURE.read_text(encoding="utf-8").splitlines())
    assert report.rows[0].alt_barcodes == ["07501068898989"]
    assert report.rows[1].alt_barcodes == []

def test_lee_metadatos_del_encabezado():
    report = parse_rows(FIXTURE.read_text(encoding="utf-8").splitlines())
    assert report.dependencia == "10 MAIPU"
    assert report.fecha == "17/06/2026"
    assert report.hora == "10:06:36"

def test_continuacion_sin_item_previo_cuenta_como_desconocida():
    report = parse_rows([CONT_07501068898989, ITEM_26726760])
    assert report.unknown_lines == 1
    assert report.rows[0].alt_barcodes == []
```

Añadir al fixture, al final, un bloque de segunda página (las mismas 8 líneas de encabezado de la Task 1 más un ítem distinto) y una prueba (ajustando `test_ignora_encabezados_y_reglas` de Task 1 a 3 ítems):

```python
def test_salta_encabezados_de_pagina_intermedios():
    report = parse_rows(FIXTURE.read_text(encoding="utf-8").splitlines())
    assert len(report.rows) == 3
    assert report.unknown_lines == 0
```

- [ ] **Step 2: Correr y ver que fallan**

Run: `python -m pytest tests/test_parser.py -v`
Expected: FAIL en las cuatro pruebas nuevas.

- [ ] **Step 3: Implementar**

Continuación: añade el código a `rows[-1].alt_barcodes`; si no hay ítem previo, cuenta como desconocida. Metadatos: busca con regex las líneas sin `|` que empiezan por `FECHA :`, `HORA  :` y `DEPENDENCIA  :` (tolerante a espacios) y quédate con lo que sigue a `:`; `fecha` es la primera `FECHA` encontrada en el archivo.

- [ ] **Step 4: Correr y ver que pasan**

Run: `python -m pytest tests/test_parser.py -v`
Expected: 7 passed.

- [ ] **Step 5: Prueba contra el archivo real (marcada `@pytest.mark.slow`)**

```python
import pytest
REAL = Path(r"C:\Users\victus\Downloads\stock17-06.xlsx")

@pytest.mark.slow
@pytest.mark.skipif(not REAL.exists(), reason="archivo real no disponible")
def test_reporte_real_completo():
    report = read_report(REAL)
    assert len(report.rows) == 26527
    assert len(report.areas()) == 69
    assert sum(len(r.alt_barcodes) for r in report.rows) == 3192
    assert any(r.plu == "23842742" for r in report.rows)  # fila desbordada a B
```

Run: `python -m pytest tests/test_parser.py -v`
Expected: 8 passed (el lento puede tardar ~10 s).

- [ ] **Step 6: Commit**

```bash
git add .
git commit -m "feat: codigos alternativos, metadatos y paginas"
```

---

### Task 3: Validación, duplicados y reporte vacío

**Files:**
- Modify: `stocky/parser.py`
- Test: `tests/test_parser.py`

**Interfaces:**
- Produces: `ParseError` con mensajes de §12; `StockReport.conflicts:int`.

- [ ] **Step 1: Escribir las pruebas que fallan**

```python
def test_reporte_vacio_lanza_error():
    with pytest.raises(ParseError, match="no trae ítems"):
        parse_rows([",".join(["x"] * 13)])

def test_plu_repetido_con_otro_codigo_se_fusiona():
    lineas = [ITEM_26726760, CONT_07501068898989, ITEM_26726760_CON_CODIGO_2]
    report = parse_rows(lineas)
    assert len(report.rows) == 1
    assert report.rows[0].barcode == "00200026726766"
    assert "00200026726799" in report.rows[0].alt_barcodes
    assert report.conflicts == 0

def test_plu_repetido_con_stock_distinto_es_conflicto():
    report = parse_rows([ITEM_26726760, ITEM_26726760_STOCK_5])
    assert len(report.rows) == 1
    assert report.rows[0].disponible == -1
    assert report.conflicts == 1

def test_reporte_de_otro_formato_lanza_error(tmp_path):
    wb = openpyxl.Workbook()
    wb.active["A1"] = "AREA"
    wb.active["B1"] = "PLU"
    p = tmp_path / "otro.xlsx"
    wb.save(p)
    with pytest.raises(ParseError, match="Formato no soportado"):
        read_report(p)
```

Definir en el test las tres constantes de línea copiando el formato del fixture, cambiando solo el código y el disponible.

- [ ] **Step 2: Correr y ver que fallan**

Run: `python -m pytest tests/test_parser.py -v`
Expected: FAIL en las tres.

- [ ] **Step 3: Implementar**

Clave de duplicado `(area, plu)`. Si coincide `desc`, `disponible`, `precio` y `costo`: fusionar uniendo el `barcode` distinto a `alt_barcodes` del primero. Si algo difiere: conservar el primero y `conflicts += 1`. Después de parsear todo, si `rows` está vacío → `ParseError("El reporte no trae ítems")`. En `read_report`, si el libro no tiene ninguna celda con `|` → `ParseError("Formato no soportado: no parece un Listado de existencias")`.

- [ ] **Step 4: Correr y ver que pasan**

Run: `python -m pytest tests/test_parser.py -v`
Expected: 11 passed.

- [ ] **Step 5: Commit**

```bash
git add .
git commit -m "feat: dedupe de PLU y reportes vacios"
```

---

### Task 4: Mapa de áreas

**Files:**
- Create: `stocky/areas.py`, `tests/test_areas.py`

**Interfaces:**
- Produces: `AREA_NAMES: dict[int,str]`; `area_name(n:int)->str|None`; `file_stem(n:int)->str` (`"09 - TELEFONIA"`, `"37"`); `header_label(n:int)->str` (`"ÁREA 09 · TELEFONIA"`, `"ÁREA 37"`).

- [ ] **Step 1: Escribir las pruebas que fallan**

```python
from stocky.areas import area_name, file_stem, header_label

def test_mapa_completo_del_letrero():
    assert area_name(9) == "TELEFONIA"
    assert area_name(99) == "TELEVISION"
    assert area_name(80) == "HORNO"
    assert len(AREA_NAMES) == 18

def test_area_fuera_del_mapa():
    assert area_name(37) is None
    assert file_stem(37) == "37"
    assert header_label(37) == "ÁREA 37"

def test_nombre_de_archivo():
    assert file_stem(9) == "09 - TELEFONIA"
    assert header_label(9) == "ÁREA 09 · TELEFONIA"
```

- [ ] **Step 2: Correr y ver que fallan**

Run: `python -m pytest tests/test_areas.py -v`
Expected: FAIL, módulo inexistente.

- [ ] **Step 3: Implementar**

Copia el mapa de §7 del spec (18 entradas: 8, 9, 10, 29, 46, 73, 74, 77, 78, 79, 80, 81, 83, 84, 85, 86, 95, 99). `file_stem` rellena a dos dígitos y solo agrega `" - NOMBRE"` si hay nombre.

- [ ] **Step 4: Correr y ver que pasan**

Run: `python -m pytest tests/test_areas.py -v`
Expected: 3 passed.

- [ ] **Step 5: Commit**

```bash
git add .
git commit -m "feat: mapa de areas del letrero"
```

---

### Task 5: Formato y presets persistentes

**Files:**
- Create: `stocky/presets.py`, `tests/test_presets.py`

**Interfaces:**
- Produces:
  - `Formato` (dataclass, slots) con `font_family="Segoe UI"`, `font_size=10`, `title_size=14`, `density="normal"`, `zebra=False`, `header_fill="gris"`, `margin_cm=1.5`, `orientation="vertical"`, `wide_desc=False`, `show_rubro=False`, `thousands=True`, `extra_header=""`, `min_stock=1`; `to_dict()->dict`, `Formato.from_dict(d:dict)->Formato`.
  - `Config` con `last_format:Formato`, `presets:dict[str,Formato]`, `last_areas:list[int]`, `last_folder:str`; `load_config(path:Path=DEFAULT_CONFIG_PATH)->Config`, `save_config(cfg:Config, path:Path=DEFAULT_CONFIG_PATH)->None`, `DEFAULT_CONFIG_PATH`.

- [ ] **Step 1: Escribir las pruebas que fallan**

```python
def test_defaults_del_formato():
    f = Formato()
    assert (f.font_family, f.font_size, f.margin_cm, f.orientation) == ("Segoe UI", 10, 1.5, "vertical")

def test_from_dict_recorta_valores_fuera_de_rango():
    f = Formato.from_dict({"font_size": 99, "margin_cm": 9, "orientation": "diagonal"})
    assert f.font_size == 16 and f.margin_cm == 2.0 and f.orientation == "vertical"

def test_ida_y_vuelta_a_json(tmp_path):
    cfg = Config()
    cfg.presets["Sala"] = Formato(font_size=12)
    cfg.last_areas = [9, 99]
    save_config(cfg, tmp_path / "config.json")
    again = load_config(tmp_path / "config.json")
    assert again.presets["Sala"].font_size == 12 and again.last_areas == [9, 99]

def test_config_corrupta_devuelve_defaults(tmp_path):
    p = tmp_path / "config.json"
    p.write_text("{ roto", encoding="utf-8")
    cfg = load_config(p)
    assert cfg.presets == {} and cfg.last_format.font_size == 10

def test_guardado_es_atomico_sin_dejar_temporales(tmp_path):
    save_config(Config(), tmp_path / "config.json")
    assert [p.name for p in tmp_path.iterdir()] == ["config.json"]
```

- [ ] **Step 2: Correr y ver que fallan**

Run: `python -m pytest tests/test_presets.py -v`
Expected: FAIL, módulo inexistente.

- [ ] **Step 3: Implementar**

`from_dict` ignora claves desconocidas, recorta `font_size` a 8–16 y `title_size` a 8–24; `margin_cm` solo acepta `{0.8, 1.0, 1.5, 2.0}` (si no, 1.5); `density` ∈ {compacta, normal, amplia}; `orientation` ∈ {vertical, horizontal}; `header_fill` ∈ {ninguna, gris, verde}. `save_config` escribe a `<path>.tmp` con `ensure_ascii=False, indent=2` y hace `Path.replace`. `load_config` captura `FileNotFoundError`, `json.JSONDecodeError` y `TypeError` → `Config()`.

- [ ] **Step 4: Correr y ver que pasan**

Run: `python -m pytest tests/test_presets.py -v`
Expected: 5 passed.

- [ ] **Step 5: Commit**

```bash
git add .
git commit -m "feat: formato configurable y presets"
```

---

### Task 6: Exportación a xlsx con formato de impresión

**Files:**
- Create: `stocky/export.py`, `tests/test_export.py`

**Interfaces:**
- Consumes: `StockReport`, `StockRow`, `Formato`, `areas.file_stem/header_label`.
- Produces: `rows_for(report, area, min_stock=1) -> list[StockRow]`; `build_workbook(rows, area, report, fmt) -> openpyxl.Workbook`; `write_area_xlsx(report, area, fmt, out_dir: Path) -> Path`; `output_dir(report, source: Path) -> Path`.

- [ ] **Step 1: Escribir las pruebas que fallan**

```python
def test_filtra_y_ordena(report_mini):
    rows = rows_for(report_mini, area=9)
    assert [r.desc for r in rows] == ["ALFA", "BETA", "GAMMA"]
    assert all(r.disponible >= 1 for r in rows)

def test_codigo_y_plu_quedan_como_texto(report_mini):
    ws = build_workbook(rows_for(report_mini, 9), 9, report_mini, Formato()).active
    codigos = [c.value for c in ws["B"][4:]]
    assert "00200026726766" in codigos

def test_configuracion_de_pagina_carta_y_una_hoja_de_ancho(report_mini):
    ws = build_workbook(rows_for(report_mini, 9), 9, report_mini, Formato(font_size=16, wide_desc=True)).active
    assert ws.page_setup.paperSize == ws.PAPERSIZE_LETTER
    assert ws.page_setup.fitToWidth == 1
    assert ws.sheet_properties.pageSetUpPr.fitToPage is True
    assert ws.print_title_rows == "1:4"

def test_pie_con_ajuste_y_totales(report_mini):
    ws = build_workbook(rows_for(report_mini, 9), 9, report_mini, Formato()).active
    assert "Página &P de &N" in ws.oddFooter.left.text
    assert "3 ítems" in ws.oddFooter.left.text or "3 ítems" in ws.oddFooter.right.text

def test_area_sin_stock_no_escribe_archivo(report_mini, tmp_path):
    assert write_area_xlsx(report_mini, area=8, fmt=Formato(), out_dir=tmp_path) is None
```

`conftest.py` define `report_mini`: tres ítems en área 9 (`ALFA` disponible 1, `BETA` disponible 0, `GAMMA` disponible 2) y uno en área 8 (`DELTA` disponible 0); el primero con un PLU con ceros y un `alt_barcodes` (`"00200026726766"` + `["07501068898989"]`).

- [ ] **Step 2: Correr y ver que fallan**

Run: `python -m pytest tests/test_export.py -v`
Expected: FAIL, módulo inexistente.

- [ ] **Step 3: Implementar**

`rows_for`: filtra `disponible >= min_stock`, ordena por `(unicodedata.normalize("NFKD", desc).encode("ascii","ignore").decode().upper(), plu)`.

`build_workbook`: fila 1 título `header_label(area)` (fuente `title_size`) más `extra_header`; fila 2 `MAIPU · fecha · hora · N ítems · U unidades` (fuente `font_size`); fila 3 vacía; fila 4 encabezados de columna (negrita, relleno según `header_fill`, borde inferior). Datos desde fila 5. Formato de número `#,##0` en DISPONIBLE y PRECIO cuando `thousands`, `0` si no. Columnas de texto con `number_format = "@"`. Anchos: PLU 11, CÓDIGO 17, DESCRIPCIÓN 44 (56 si `wide_desc`), DISPONIBLE 11, PRECIO 13. `alt_barcodes` van en una fila extra debajo del ítem, solo columna CÓDIGO, fuente `font_size - 2`, sin relleno de zebra. `density` fija alto de fila (compacta 12, normal 15, amplia 18). Impresión: `paperSize = PAPERSIZE_LETTER`, `orientation` según `fmt`, márgenes en pulgadas (`cm / 2.54`), `fitToPage=True`, `fitToWidth=1`, `fitToHeight=0`, `print_title_rows="1:4"`, `print_area` hasta la última fila/columna, footer `&L N ítems · U unidades   &R Página &P de &N`, `ws.sheet_view.showGridLines = False`.

`write_area_xlsx`: si `rows_for` viene vacío devuelve `None`; si no, crea `out_dir`, arma el libro, guarda en `out_dir / f"{file_stem(area)}.xlsx"` y devuelve la ruta.

`output_dir(report, source)`: `source.parent / f"Stock {report.fecha.replace('/', '-')}"` invertido a `YYYY-MM-DD` (`17/06/2026` → `2026-06-17`).

- [ ] **Step 4: Correr y ver que pasan**

Run: `python -m pytest tests/test_export.py -v`
Expected: 5 passed.

- [ ] **Step 5: Commit**

```bash
git add .
git commit -m "feat: exportacion xlsx lista para imprimir"
```

---

### Task 7: PDF con Excel y envío a impresora

**Files:**
- Create: `stocky/pdf.py`, `tests/test_pdf.py`

**Interfaces:**
- Produces: `excel_available() -> bool`; `export_pdf(xlsx: Path, pdf: Path) -> Path`; `print_pdf(pdf: Path) -> None`; `PdfError(Exception)`.

- [ ] **Step 1: Escribir las pruebas que fallan**

```python
@pytest.mark.skipif(not excel_available(), reason="sin Excel")
def test_exporta_pdf_de_una_pagina(xlsx_mini, tmp_path):
    out = export_pdf(xlsx_mini, tmp_path / "a.pdf")
    assert out.exists() and out.stat().st_size > 1000
    import pypdfium2 as pdfium
    assert len(pdfium.PdfDocument(str(out))) == 1

def test_error_si_el_xlsx_no_existe(tmp_path):
    with pytest.raises(PdfError, match="no existe"):
        export_pdf(tmp_path / "nada.xlsx", tmp_path / "a.pdf")
```

`xlsx_mini` es una fixture que usa `build_workbook` con `report_mini` y guarda el archivo en `tmp_path`.

- [ ] **Step 2: Correr y ver que fallan**

Run: `python -m pytest tests/test_pdf.py -v`
Expected: FAIL, módulo inexistente.

- [ ] **Step 3: Implementar**

`export_pdf`: valida existencia; `win32com.client.DispatchEx("Excel.Application")` con `Visible=False`, `DisplayAlerts=False`; abre el xlsx con `ReadOnly=True`, `ExportAsFixedFormat(0, str(pdf.resolve()))`, cierra sin guardar; `Quit()` en `finally`. Envuelve `com_error` en `PdfError` con el texto de Excel. `excel_available` intenta crear y cerrar la app y devuelve `False` ante cualquier excepción. `print_pdf` usa `os.startfile(str(pdf), "print")`.

- [ ] **Step 4: Correr y ver que pasan**

Run: `python -m pytest tests/test_pdf.py -v`
Expected: 2 passed (o 1 passed + 1 skipped sin Excel).

- [ ] **Step 5: Commit**

```bash
git add .
git commit -m "feat: exportacion a PDF y envio a impresora"
```

---

### Task 8: Render de páginas para la vista previa

**Files:**
- Create: `stocky/preview.py`, `tests/test_preview.py`

**Interfaces:**
- Produces: `render_pages(pdf: Path, scale: float = 1.5) -> list[bytes]` (PNG por página); `clear_cache() -> None`.

- [ ] **Step 1: Escribir las pruebas que fallan**

```python
def test_renderiza_una_pagina_como_png(pdf_mini):
    pages = render_pages(pdf_mini)
    assert len(pages) == 1
    assert pages[0][:8] == b"\x89PNG\r\n\x1a\n"

def test_cache_no_relee_el_mismo_pdf(pdf_mini, monkeypatch):
    clear_cache()
    primera = render_pages(pdf_mini)
    monkeypatch.setattr(preview, "_open_pdf", lambda p: (_ for _ in ()).throw(AssertionError("releyó")))
    assert render_pages(pdf_mini) == primera

def test_escala_mayor_da_mas_pixeles(pdf_mini):
    from PIL import Image
    import io
    chica = Image.open(io.BytesIO(render_pages(pdf_mini, scale=1.0)[0])).size
    grande = Image.open(io.BytesIO(render_pages(pdf_mini, scale=2.0)[0])).size
    assert grande[0] > chica[0]
```

- [ ] **Step 2: Correr y ver que fallan**

Run: `python -m pytest tests/test_preview.py -v`
Expected: FAIL, módulo inexistente.

- [ ] **Step 3: Implementar**

`_open_pdf(path)` abre `pypdfium2.PdfDocument`. Cada página se renderiza a bitmap RGB y se codifica a PNG con Pillow (`Image.frombuffer("RGB", (w, h), buf, "raw", "BGR", 0, 1)` o `page.render(...).to_pil()`), devolviendo bytes. La caché es un `dict[(str(pdf), mtime_ns, scale), list[bytes]]` con máximo 20 entradas (FIFO).

- [ ] **Step 4: Correr y ver que pasan**

Run: `python -m pytest tests/test_preview.py -v`
Expected: 3 passed.

- [ ] **Step 5: Commit**

```bash
git add .
git commit -m "feat: render de paginas para vista previa"
```

---

### Task 9: Lógica de la UI sin ventana

**Files:**
- Create: `stocky/ui_state.py`, `tests/test_ui_state.py`

**Interfaces:**
- Produces: `DEFAULT_AREAS = [8, 9, 29, 46, 84, 99]`; `initial_selection(disponibles: list[int], last: list[int]) -> list[int]`; `status_text(report: StockReport, selection: list[int]) -> str`; `resolve_font(family: str, available: list[str]) -> str`.

- [ ] **Step 1: Escribir las pruebas que fallan**

```python
def test_seleccion_inicial_usa_las_seis_de_siempre(report_mini):
    assert initial_selection(sorted(report_mini.areas()), []) == [9]  # solo 9 tiene stock en el mini

def test_seleccion_inicial_respeta_la_ultima(report_mini):
    assert initial_selection([8, 9, 29], [29]) == [29]

def test_seleccion_ignora_areas_que_no_existen(report_mini):
    assert initial_selection([9], [9, 99]) == [9]

def test_status_resume_anomalias(report_mini):
    report_mini.unknown_lines = 2
    report_mini.conflicts = 1
    texto = status_text(report_mini, [9])
    assert "2 líneas no reconocidas" in texto and "1 duplicado conflictivo" in texto

def test_fuente_inexistente_cae_en_segoe_ui():
    assert resolve_font("NoExiste", ["Segoe UI", "Arial"]) == "Segoe UI"
    assert resolve_font("Arial", ["Segoe UI", "Arial"]) == "Arial"
```

- [ ] **Step 2: Correr y ver que fallan**

Run: `python -m pytest tests/test_ui_state.py -v`
Expected: FAIL, módulo inexistente.

- [ ] **Step 3: Implementar**

`initial_selection` devuelve la intersección de `last` con las áreas con stock; si `last` está vacía, devuelve la intersección con `DEFAULT_AREAS`; si el resultado queda vacío, devuelve todas las áreas con stock. `status_text` arma `"26.527 ítems · 6 áreas · 2 líneas no reconocidas · 1 duplicado conflictivo"`, omitiendo las partes en cero. `resolve_font` devuelve la familia si está en `available`, si no `"Segoe UI"`.

- [ ] **Step 4: Correr y ver que pasan**

Run: `python -m pytest tests/test_ui_state.py -v`
Expected: 5 passed.

- [ ] **Step 5: Commit**

```bash
git add .
git commit -m "feat: logica de seleccion y estado"
```

---

### Task 10: Ventana con panel lateral y vista previa

**Files:**
- Create: `stocky/theme.py`, `stocky/ui.py`, `tests/test_ui_smoke.py`

**Interfaces:**
- Consumes: todo lo anterior.
- Produces: `stocky.ui.run(report: StockReport, source: Path) -> None`; `stocky.ui.App(report: StockReport, source: Path, root: tk.Tk)` con atributos `selection: list[int]`, `current_area: int`, `root`; `stocky.theme.COLORS: dict[str,str]`, `stocky.theme.FONT_STACK: list[str]`.

- [ ] **Step 1: Escribir la prueba que falla (smoke, sin interacción)**

```python
def test_la_ventana_se_construye_con_un_reporte(report_mini, tmp_path, monkeypatch):
    monkeypatch.setenv("APPDATA", str(tmp_path))
    app = App(report_mini, tmp_path / "stock17-06.xlsx", root=tk.Tk())
    assert app.selection == [9]
    assert app.current_area == 9
    app.root.destroy()
```

- [ ] **Step 2: Correr y ver que falla**

Run: `python -m pytest tests/test_ui_smoke.py -v`
Expected: FAIL, `App` no existe.

- [ ] **Step 3: Implementar `theme.py`**

`COLORS = {"bg": "#F8FAFC", "panel": "#FFFFFF", "ink": "#0F172A", "muted": "#64748B", "line": "#E6E8EA", "slate": "#334155", "accent": "#059669", "accent_ink": "#FFFFFF", "preview_bg": "#EEF1F4"}`. `FONT_STACK = ["Segoe UI", "Calibri", "Arial", "Tahoma", "Verdana", "Consolas"]`. Estilos ttk: `Sidebar.TFrame`, `H1.TLabel`, `Muted.TLabel`, `Accent.TButton`, `Ghost.TButton`, `Area.TCheckbutton`.

- [ ] **Step 4: Implementar `ui.py`**

Estructura: ventana 1180×760, título `Stocky — <fecha>`. Barra superior con nombre del archivo, `status_text` y botón "Abrir carpeta". Panel izquierdo (260 px, scrollable): selector de fuente (las familias de `tkinter.font.families()` resueltas con `resolve_font`), tamaño del cuerpo (Spinbox 8–16), tamaño del título (Spinbox 8–24), densidad (Combobox), filas alternadas, color de encabezado, márgenes (Combobox con las 4 opciones, mostradas en cm), orientación, descripción ancha, mostrar rubro, separador de miles, encabezado extra (Entry), lista de áreas con `Checkbutton` y conteo, botones "Todas", "Las de siempre", "Guardar como preset", "Presets ▾", "Restablecer". Zona derecha: pestañas por área seleccionada con Notebook, y un `Canvas` con scroll vertical que muestra las páginas como `PhotoImage` (base64 de los bytes de `render_pages`). Barra inferior: estado del render y botones "Generar todos", "Imprimir", "Abrir PDF".

Comportamiento: cualquier cambio de control reprograma `root.after(400, self._refresh_preview)`; `_refresh_preview` marca el área actual, escribe el xlsx en la carpeta temporal (`tempfile.mkdtemp(prefix="stocky_")`), llama `export_pdf` y `render_pages`, y pinta. Si `export_pdf` falla, muestra el mensaje en la barra inferior y deja los xlsx disponibles. "Generar todos" escribe en `output_dir` cada área marcada (xlsx y pdf) y abre la carpeta con `os.startfile`; si la escritura falla por archivo abierto o permiso, muestra la ruta y el motivo en la barra inferior y no corta el resto. "Imprimir" llama `print_pdf` sobre el pdf del área visible. Al cerrar, guarda `Config` con formato, áreas y carpeta.

- [ ] **Step 5: Correr y ver que pasan**

Run: `python -m pytest tests/test_ui_smoke.py -v`
Expected: 1 passed.

- [ ] **Step 6: Commit**

```bash
git add .
git commit -m "feat: ventana con panel de formato y vista previa"
```

---

### Task 11: Entrada por arrastre, ícono y cierre

**Files:**
- Create: `stocky/app.py`, `herramientas/crear_acceso.py`, `abrir consola.cmd`, `tests/test_app.py`

**Interfaces:**
- Produces: `stocky.app.main(argv: list[str]) -> int`.

- [ ] **Step 1: Escribir las pruebas que fallan**

```python
def test_sin_argumentos_explica_el_uso(capsys):
    assert main([]) == 2
    assert "arrastra" in capsys.readouterr().err.lower()

def test_archivo_inexistente_devuelve_error(capsys, tmp_path):
    assert main([str(tmp_path / "no.xlsx")]) == 2
    assert "no existe" in capsys.readouterr().err.lower()

def test_reporte_vacio_muestra_el_mensaje(monkeypatch, capsys, tmp_path):
    monkeypatch.setattr(app, "read_report", lambda p: (_ for _ in ()).throw(ParseError("El reporte no trae ítems")))
    assert main([str(tmp_path)]) == 1
    assert "no trae ítems" in capsys.readouterr().err
```

- [ ] **Step 2: Correr y ver que fallan**

Run: `python -m pytest tests/test_app.py -v`
Expected: FAIL, módulo inexistente.

- [ ] **Step 3: Implementar `app.py`**

`main` valida argumentos, llama `read_report`, y en éxito `ui.run(report, path)`. Errores → mensaje en `stderr` (y `messagebox.showerror` cuando hay consola oculta) y código de salida 1 (datos) o 2 (uso). `if __name__ == "__main__": raise SystemExit(main(sys.argv[1:]))`.

- [ ] **Step 4: Correr y ver que pasan**

Run: `python -m pytest tests/test_app.py -v`
Expected: 3 passed.

- [ ] **Step 5: Crear lanzadores**

`herramientas/crear_acceso.py` genera `assets/stocky.ico` (un ícono simple con Pillow) y crea `Stocky.lnk` en la raíz del proyecto con `win32com.client.Dispatch("WScript.Shell")`: target `sys.executable` reemplazado por `pythonw.exe`, argumentos `"<ruta>\stocky\app.py"`, icono `<ruta>\assets\stocky.ico`, `WorkingDirectory` el del proyecto. `abrir consola.cmd`: `@echo off` + `python "%~dp0stocky\app.py" %*` + `pause`.

- [ ] **Step 6: Probar el arrastre de verdad**

Arrastrar `C:\Users\victus\Downloads\stock17-06.xlsx` sobre `Stocky.lnk`.
Expected: abre el panel con 6 áreas marcadas, vista previa de TELEFONÍA, y sin errores en la barra inferior.

- [ ] **Step 7: Commit**

```bash
git add .
git commit -m "feat: entrada por arrastre y acceso directo"
```

---

### Task 12: Verificación final con el archivo real

**Files:**
- Create: `docs/superpowers/plans/verificacion-2026-10-22.md` (registro de la corrida)

- [ ] **Step 1: Correr toda la suite**

Run: `python -m pytest -v`
Expected: todo verde; los tests `slow` y de Excel corren, no se saltan.

- [ ] **Step 2: Generar las seis áreas del reporte real**

Abrir el panel con `stock17-06.xlsx`, dejar marcadas 8, 9, 29, 46, 84, 99, "Generar todos".
Expected: `Stock 2026-06-17\` con 12 archivos; ningún área sin stock entre las seis.

- [ ] **Step 3: Verificar cifras contra el reporte**

Para el área 9: contar en el xlsx generado las filas de datos y la suma de DISPONIBLE; comparar con el parser (`rows_for(report, 9)`).
Expected: coinciden exactamente; el pie impreso muestra el mismo número.

- [ ] **Step 4: Verificar el ajuste a una hoja de ancho**

Abrir el PDF del área 9 y del área 99 en un visor, comprobar que ninguna página parte columnas.
Expected: todas las páginas muestran las 5 columnas completas.

- [ ] **Step 5: Probar el peor caso de formato**

Con el área 9 en pantalla: Segoe UI 16, descripción ancha, márgenes 2,0 cm.
Expected: la vista previa sigue entrando a lo ancho (escala reducida); al imprimir, igual.

- [ ] **Step 6: Probar los caminos de error**

Arrastrar `stock19-10.xlsx` (vacío) → mensaje "El reporte no trae ítems", sin carpeta nueva.
Renombrar el `config.json` a basura y abrir la app → arranca con defaults.
Expected: ambos casos se comportan como dice el spec.

- [ ] **Step 7: Registrar y commitear**

Anotar en `docs/superpowers/plans/verificacion-2026-10-22.md` los conteos, páginas por área y cualquier diferencia encontrada. Commit final.
