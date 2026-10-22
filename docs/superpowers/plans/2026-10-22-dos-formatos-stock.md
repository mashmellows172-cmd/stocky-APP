# Soporte de dos formatos de stock — Plan de implementación

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Que Stocky abra automáticamente el reporte histórico con pipes y el reporte habitual en columnas, sin cambiar la vista, la exportación ni la impresión.

**Architecture:** `read_report` lee la hoja una vez y detecta el formato. Si encuentra el encabezado de tabla (`PLU` + `DESCRIPCION PRODUCTO`) usa un parser nuevo que recorre las columnas; si no, conserva el parser de pipes actual. Ambos producen el mismo `StockReport`, así que el resto de la app no cambia.

**Tech Stack:** Python 3.14, openpyxl, pytest, Tkinter (sin cambios).

**Spec:** `docs/superpowers/specs/2026-10-22-dos-formatos-stock-design.md`

## Global Constraints

- El formato de pipes de junio debe seguir funcionando exactamente igual: `tests/test_parser.py::test_reporte_real_completo` debe seguir verde.
- No agregar dependencias nuevas.
- No tocar la interfaz, la exportación, los presets ni la impresión.
- Mantener los nombres y tipos de `StockReport` y `StockRow`.
- Los campos que el formato de columnas no trae (`un_emp`, `un_despa`, `existencia`, `costo`) valen `0`.
- `disponible` y `precio` vacíos se interpretan como `0`.
- Windows: ejecutar siempre `python -m pytest -q` desde la raíz del repo.
- Commit al final de cada tarea.

## Review Focus

1. **Encabezados repetidos y filas de ruido** en medio de la tabla: deben ignorarse, no crear ítems ni inflar `unknown_lines`.
2. **PLU repetido con otro código y los mismos datos**: debe fusionarse en una fila con `alt_barcodes`, sin duplicar el ítem.
3. **Disponible o precio vacíos**: no deben romper la lectura; el ítem queda en 0 y el filtro `disponible >= 1` lo excluye.
4. **Un `|` dentro de una descripción**: el archivo debe seguir detectándose como tabla, no como formato de pipes.
5. **Tabla sin encabezado o archivo desconocido**: no adivinar; debe seguir apareciendo el error visible actual.

---

### Task 1: Detección automática y parser de columnas

**Files:**
- Modify: `stocky/parser.py`
- Test: `tests/test_parser.py`

**Interfaces:**
- Consumes: `StockReport`, `StockRow`, `_read_metadata`, `_dedupe`, `_last_int` (ya existen en `stocky/parser.py`).
- Produces:
  - `parse_table_rows(rows: Iterable[Sequence[object]]) -> StockReport`
  - `_looks_like_table(rows: Iterable[Sequence[object]]) -> bool`
  - `read_report(path: Path) -> StockReport` (misma firma y mismo comportamiento para pipes).

- [ ] **Step 1: Escribir las pruebas que fallan**

Agregar al final de `tests/test_parser.py` el helper y las pruebas siguientes. El helper imita el archivo real: metadatos en la columna D, encabezado en celdas separadas, una fila de ruido con solo un código y un encabezado repetido en medio de los datos.

```python
def _write_column_report(path: Path) -> Path:
    import openpyxl

    wb = openpyxl.Workbook()
    ws = wb.active
    filas = [
        [None, None, None, "FECHA : 19/10/2026"],
        [None, None, None, "HORA  : 10:14:33"],
        [None, None, None, "DEPENDENCIA   :   10 MAIPU"],
        [None, "AREA", "NIVEL DE AGREGACION", "PLU", "CODIGO DE BARRA", "DESCRIPCION PRODUCTO", "DISPONIBLE", "PRECIO DE"],
        [None, 9, "7565 TELEFONIA", "26726760", "00200026726766", "CELULAR ALFA", 2, 59990],
        [None, 9, "7565 TELEFONIA", "26726760", "07501068898989", "CELULAR ALFA", 2, 59990],
        [None, None, None, None, "07501068898989", None, None, None],
        [None, "AREA", "NIVEL DE AGREGACION", "PLU", "CODIGO DE BARRA", "DESCRIPCION PRODUCTO", "DISPONIBLE", "PRECIO DE"],
        [None, 9, "7565 TELEFONIA", "26726761", "00200026726799", "CELULAR BETA", 0, 37990],
        [None, 8, "7881 AUDIO", "26726762", "00200026726762", "PARLANTE | GAMMA", 1, 15000],
        [None, 10, "9999 VARIOS", "26726764", "00200026726764", "REPUESTO", None, None],
    ]
    for fila in filas:
        ws.append(fila)
    wb.save(path)
    return path


def test_lee_reporte_en_columnas(tmp_path):
    p = _write_column_report(tmp_path / "columnas.xlsx")
    report = read_report(p)
    assert (report.dependencia, report.fecha, report.hora) == ("10 MAIPU", "19/10/2026", "10:14:33")
    assert len(report.rows) == 4
    assert report.unknown_lines == 0
    assert report.conflicts == 0
    alfa = report.rows[0]
    assert (alfa.area, alfa.rubro, alfa.plu, alfa.barcode, alfa.desc) == (
        9,
        "7565 TELEFONIA",
        "26726760",
        "00200026726766",
        "CELULAR ALFA",
    )
    assert (alfa.un_emp, alfa.un_despa, alfa.existencia, alfa.disponible, alfa.precio, alfa.costo) == (0, 0, 0, 2, 59990, 0)


def test_columnas_ignoran_encabezados_y_ruido(tmp_path):
    p = _write_column_report(tmp_path / "columnas.xlsx")
    report = read_report(p)
    assert len(report.rows) == 4
    assert report.unknown_lines == 0
    assert [fila.desc for fila in report.rows] == ["CELULAR ALFA", "CELULAR BETA", "PARLANTE | GAMMA", "REPUESTO"]


def test_columnas_fusionan_plu_repetido(tmp_path):
    p = _write_column_report(tmp_path / "columnas.xlsx")
    report = read_report(p)
    alfa = report.rows[0]
    assert alfa.plu == "26726760"
    assert alfa.alt_barcodes == ["07501068898989"]
    assert len([fila for fila in report.rows if fila.plu == "26726760"]) == 1


def test_columnas_con_campos_vacios_quedan_en_cero(tmp_path):
    p = _write_column_report(tmp_path / "columnas.xlsx")
    report = read_report(p)
    repuesto = report.rows[3]
    assert (repuesto.disponible, repuesto.precio) == (0, 0)
    assert report.areas() == {8: 1, 9: 1}


def test_columnas_con_pipe_en_descripcion_no_se_confunde(tmp_path):
    p = _write_column_report(tmp_path / "columnas.xlsx")
    report = read_report(p)
    assert report.rows[2].desc == "PARLANTE | GAMMA"


def test_tabla_sin_encabezado_no_se_adivina(tmp_path):
    import openpyxl

    wb = openpyxl.Workbook()
    wb.active.append([None, 9, "7565 TELEFONIA", "26726760", "00200026726766", "CELULAR ALFA", 2, 59990])
    p = tmp_path / "sin_encabezado.xlsx"
    wb.save(p)
    with pytest.raises(ParseError, match="Formato no soportado"):
        read_report(p)
```

- [ ] **Step 2: Correr las pruebas para ver que fallan**

Run: `python -m pytest tests/test_parser.py -q`

Expected: las seis pruebas nuevas fallan. Las que leen el archivo de columnas deben fallar con `ParseError: Formato no soportado`; `test_tabla_sin_encabezado_no_se_adivina` puede pasar ya.

- [ ] **Step 3: Implementar la detección y el parser**

En `stocky/parser.py`:

1. Agregar `Sequence` al import de `typing`.
2. Agregar `_looks_like_table(rows: Iterable[Sequence[object]]) -> bool`: devuelve `True` si alguna fila tiene, entre sus celdas recortadas, `PLU` y una celda que empiece con `DESCRIPCION PRODUCTO`.
3. Agregar `parse_table_rows(rows: Iterable[Sequence[object]]) -> StockReport`:
   - Recorrer las filas y normalizar cada celda a texto recortado (`""` si es `None`), completando hasta 8 columnas.
   - Llamar `_read_metadata("".join(values), report)` en cada fila.
   - Aceptar como ítem solo si: `values[1]` es numérico, `values[3]` es numérico de 5 a 8 dígitos, `values[4]` no está vacío y `values[5]` no está vacío.
   - Crear `StockRow(area=int(values[1]), rubro=values[2], plu=values[3], barcode=values[4], desc=values[5], un_emp=0, un_despa=0, existencia=0, disponible=_last_int(values[6]), precio=_last_int(values[7]), costo=0)`.
   - Ignorar en silencio encabezados, filas vacías y códigos sueltos.
   - Aplicar `_dedupe(report.rows)` y lanzar `ParseError("El reporte no trae ítems")` si no quedó ninguna fila.
4. Refactorizar `read_report` para leer la hoja una sola vez:
   - `rows = list(sheet.iter_rows(values_only=True))` antes de cerrar el libro.
   - Si `_looks_like_table(rows)`, usar `parse_table_rows(rows)`.
   - Si no, construir `lines` como hoy y exigir `|`; si no hay pipes, mantener el `ParseError("Formato no soportado: no parece un Listado de existencias")`; si hay, usar `parse_rows(lines)`.
   - Asignar `report.source = path` y devolver el reporte.

La detección de tabla debe ir primero: un `|` dentro de una descripción no debe desviar el archivo al parser de pipes.

- [ ] **Step 4: Correr las pruebas nuevas**

Run: `python -m pytest tests/test_parser.py -q`

Expected: todas las pruebas de `test_parser.py` pasan, incluyendo las existentes.

- [ ] **Step 5: Correr la suite completa**

Run: `python -m pytest -q`

Expected: toda la suite pasa; el test lento de junio sigue verde.

- [ ] **Step 6: Commit**

```bash
git add stocky/parser.py tests/test_parser.py
git commit -m "feat: lee reportes de stock en formato de columnas"
```

---

### Task 2: Verificación con el archivo real y documentación

**Files:**
- Modify: `tests/test_parser.py`
- Modify: `docs/superpowers/plans/verificacion-2026-10-22.md`
- Modify: `.superpowers/sdd/2026-10-22-stocky-panel/progress.md` (ignorado por git)

**Interfaces:**
- Consumes: `read_report(path: Path) -> StockReport` de la tarea 1.
- Produces: evidencia de que el archivo real de octubre se lee completo y la documentación actualizada.

- [ ] **Step 1: Agregar la prueba lenta del archivo real**

Agregar al final de `tests/test_parser.py`:

```python
REAL_OCTUBRE = Path(r"C:\Users\victus\Downloads\stock19-10.xlsx")


@pytest.mark.slow
@pytest.mark.skipif(not REAL_OCTUBRE.exists(), reason="archivo real no disponible")
def test_reporte_de_columnas_real():
    report = read_report(REAL_OCTUBRE)
    assert len(report.rows) == 30424
    assert report.conflicts == 0
    assert report.unknown_lines == 0
    assert len({row.area for row in report.rows}) == 66
    assert len(report.areas()) == 66
    assert report.areas()[8] == 39
    assert report.areas()[9] == 107
    assert report.areas()[29] == 133
    assert report.areas()[46] == 36
    assert report.areas()[84] == 17
    assert report.areas()[99] == 74
    assert sum(row.disponible for row in report.rows) == 299342
    assert (report.dependencia, report.fecha, report.hora) == ("10 MAIPU", "19/10/2026", "10:14:33")
    primero = report.rows[0]
    assert (primero.area, primero.plu, primero.desc, primero.disponible, primero.precio) == (
        98,
        "25985176",
        "BABY MINK BM322 BEBE.RAYAS 9/12MESES",
        1,
        4450,
    )
```

- [ ] **Step 2: Correr la prueba real**

Run: `python -m pytest tests/test_parser.py::test_reporte_de_columnas_real -q`

Expected: `1 passed`. Si falla por una diferencia de conteo, corregir la lectura del parser; no ajustar los números del test sin volver a contar el archivo real.

- [ ] **Step 3: Correr toda la suite**

Run: `python -m pytest -q`

Expected: toda la suite pasa, incluyendo los dos tests lentos (junio y octubre).

- [ ] **Step 4: Actualizar la verificación**

En `docs/superpowers/plans/verificacion-2026-10-22.md`:

- Agregar una sección “Segundo formato: reporte en columnas” con:
  - 30.424 ítems únicos, 66 áreas, 0 conflictos y 299.342 unidades.
  - Las seis áreas: 08 → 39 ítems con stock, 09 → 107, 29 → 133, 46 → 36, 84 → 17, 99 → 74.
  - Metadatos: 10 MAIPU, 19/10/2026, 10:14:33.
  - Que el archivo se detecta solo, sin elegir formato.
- Corregir el bullet de “Caminos de error” que hoy dice que `stock19-10.xlsx` da “Formato no soportado”: ese archivo ahora se lee correctamente. Reemplazarlo por un archivo realmente desconocido o corrupto como ejemplo del error.

En `.superpowers/sdd/2026-10-22-stocky-panel/progress.md`, agregar una línea de ruling:

- “Se agregó un segundo parser para el formato de columnas; se detecta por el encabezado `PLU` + `DESCRIPCION PRODUCTO`; el formato de pipes queda intacto. Coste: dos caminos de lectura que mantener.”

- [ ] **Step 5: Revisar el diff y commitear**

```bash
git diff --check
git add stocky/parser.py tests/test_parser.py docs/superpowers/plans/verificacion-2026-10-22.md
git commit -m "test: verifica el reporte real en columnas"
```

Expected: el commit incluye la prueba real y la documentación; el ledger queda fuera por estar ignorado.
