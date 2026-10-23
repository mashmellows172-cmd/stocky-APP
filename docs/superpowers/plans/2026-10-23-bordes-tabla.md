# Bordes de la tabla — Plan de implementación

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Agregar dos selectores independientes para dibujar líneas alrededor de cada celda y alrededor de toda la tabla, sin cambiar los márgenes de hoja ni el aspecto por defecto.

**Architecture:** `Formato` gana dos campos persistidos. La exportación aplica bordes de openpyxl sobre la grilla del encabezado y los datos, y después pisa el perímetro con el estilo del contorno. El panel agrega dos combobox que actualizan el formato y disparan la vista previa existente.

**Tech Stack:** Python 3.14, openpyxl, Tkinter/ttk, pytest.

**Spec:** `docs/superpowers/specs/2026-10-22-bordes-tabla-design.md`

## Global Constraints

- Opciones de línea: `ninguno`, `fino`, `medio`, `grueso`, `punteado`.
- Defaults: `cell_border="ninguno"`, `table_border="ninguno"`; la salida actual no debe cambiar.
- `cell_border` cubre encabezado, datos y filas de códigos alternativos.
- `table_border` cubre solo el perímetro de encabezado + datos; no título ni metadatos.
- El perímetro pisa el estilo de celda en sus lados exteriores.
- Si `cell_border` es `ninguno`, se conserva la línea fina actual debajo del encabezado.
- No tocar la vista, exportación/impresión ni los márgenes de hoja.
- Windows: correr `python -m pytest -q` desde la raíz.

## Review Focus

1. **Filas de códigos alternativos:** deben recibir borde en todas sus columnas para que la grilla no se corte.
2. **Estilos independientes:** con celda fina y contorno grueso, el perímetro debe quedar grueso y los separadores internos finos.
3. **Compatibilidad:** configuraciones y presets viejos, sin los campos nuevos, deben seguir abriendo igual.
4. **Estilos de openpyxl:** `punteado`/`grueso` deben mapear a estilos válidos y no romper la exportación.
5. **Alcance:** título, metadatos y márgenes de hoja no deben recibir bordes por accidente.

---

### Task 1: Modelo y bordes en la exportación

**Files:**
- Modify: `stocky/presets.py`
- Modify: `stocky/export.py`
- Test: `tests/test_presets.py`
- Test: `tests/test_export.py`

**Interfaces:**
- Consumes: `Formato`, `build_workbook`.
- Produces:
  - `Formato.cell_border: str = "ninguno"`
  - `Formato.table_border: str = "ninguno"`
  - `BORDER_STYLES = ("ninguno", "fino", "medio", "grueso", "punteado")` en `presets.py`.
  - `_apply_table_borders(ws, fmt, first_row: int, last_row: int, last_column: int) -> None` en `export.py`.

- [ ] **Step 1: Escribir las pruebas que fallan**

Agregar a `tests/test_presets.py`:

```python
def test_bordes_por_defecto_y_saneamiento():
    f = Formato()
    assert (f.cell_border, f.table_border) == ("ninguno", "ninguno")
    malo = Formato.from_dict({"cell_border": "doble", "table_border": "grueso"})
    assert (malo.cell_border, malo.table_border) == ("ninguno", "grueso")


def test_bordes_ida_y_vuelta(tmp_path):
    cfg = Config()
    cfg.last_format = Formato(cell_border="punteado", table_border="grueso")
    save_config(cfg, tmp_path / "config.json")
    again = load_config(tmp_path / "config.json")
    assert (again.last_format.cell_border, again.last_format.table_border) == ("punteado", "grueso")
```

Agregar a `tests/test_export.py` (con `import pytest`):

```python
def test_sin_bordes_conserva_la_salida_actual(report_mini):
    ws = build_workbook(rows_for(report_mini, 9), 9, report_mini, Formato()).active
    assert ws["A4"].border.bottom.style == "thin"
    assert ws["A5"].border.left.style is None
    assert ws["A5"].border.bottom.style is None


@pytest.mark.parametrize(
    ("estilo", "openpyxl_style"),
    [("fino", "thin"), ("medio", "medium"), ("grueso", "thick"), ("punteado", "dotted")],
)
def test_estilos_de_celda_y_contorno(report_mini, estilo, openpyxl_style):
    fmt = Formato(cell_border=estilo, table_border=estilo)
    ws = build_workbook(rows_for(report_mini, 9), 9, report_mini, fmt).active
    assert ws["A4"].border.top.style == openpyxl_style
    assert ws["A5"].border.left.style == openpyxl_style
    assert ws["B5"].border.left.style == openpyxl_style
    assert ws.cell(ws.max_row, ws.max_column).border.bottom.style == openpyxl_style


def test_celda_y_contorno_independientes(report_mini):
    fmt = Formato(cell_border="fino", table_border="grueso")
    ws = build_workbook(rows_for(report_mini, 9), 9, report_mini, fmt).active
    assert ws["A4"].border.left.style == "thick"
    assert ws["A4"].border.top.style == "thick"
    assert ws["A4"].border.bottom.style == "thin"
    assert ws["B4"].border.left.style == "thin"
    assert ws["E4"].border.right.style == "thick"
    assert ws["A5"].border.left.style == "thick"
    assert ws["B5"].border.left.style == "thin"
    assert ws["B5"].border.right.style == "thin"


def test_codigo_alternativo_tambien_lleva_grilla(report_mini):
    fmt = Formato(cell_border="fino", table_border="grueso")
    ws = build_workbook(rows_for(report_mini, 9), 9, report_mini, fmt).active
    # ALFA es la primera fila y su código alternativo queda en la fila 6.
    assert ws["A6"].border.left.style == "thick"
    assert ws["B6"].border.right.style == "thin"
    assert ws["E6"].border.right.style == "thick"
```

- [ ] **Step 2: Correr las pruebas para ver que fallan**

Run: `python -m pytest tests/test_presets.py tests/test_export.py -q`
Expected: fallan por `TypeError` de campos nuevos o por bordes ausentes.

- [ ] **Step 3: Implementar el modelo**

En `stocky/presets.py`:

- Agregar `BORDER_STYLES = ("ninguno", "fino", "medio", "grueso", "punteado")`.
- Agregar a `Formato`: `cell_border: str = "ninguno"` y `table_border: str = "ninguno"`.
- En `_sanitize`, para esos dos campos devolver el valor solo si está en `BORDER_STYLES`.

- [ ] **Step 4: Implementar los bordes**

En `stocky/export.py`:

- Agregar `_BORDER_SIDES = {"ninguno": None, "fino": "thin", "medio": "medium", "grueso": "thick", "punteado": "dotted"}` y `BORDER_COLOR = "808080"`.
- Agregar `_border_side(style: str)` que devuelve `Side(style=..., color=BORDER_COLOR)` o `None`.
- Agregar `_apply_table_borders(...)` con la firma indicada: recorrer todas las celdas entre `HEADER_ROW` y `last_row`, columnas 1..`last_column`; usar `cell_side` en los cuatro lados; reemplazar los lados exteriores por `table_side` cuando correspondan. Si es la fila del encabezado y no hay borde de celda, conservar `THIN` en el borde inferior.
- Llamarlo después de calcular `last_row`, antes de configurar el área de impresión. Quitar la asignación individual `cell.border = Border(bottom=THIN)` del encabezado, porque el helper ya la aplica.

- [ ] **Step 5: Correr las pruebas y la suite**

Run: `python -m pytest tests/test_presets.py tests/test_export.py -q`
Expected: todos pasan.

Run: `python -m pytest -q`
Expected: toda la suite pasa.

- [ ] **Step 6: Commit**

```bash
git add stocky/presets.py stocky/export.py tests/test_presets.py tests/test_export.py
git commit -m "feat: agrega bordes configurables a la tabla"
```

---

### Task 2: Controles del panel y verificación final

**Files:**
- Modify: `stocky/ui.py`
- Test: `tests/test_ui_smoke.py`
- Modify: `docs/superpowers/plans/verificacion-2026-10-22.md`

**Interfaces:**
- Consumes: `Formato.cell_border`, `Formato.table_border`, `BORDER_STYLES`.
- Produces: `v_cell_border` y `v_table_border` en `App`; labels `Ninguno`, `Fino`, `Medio`, `Grueso`, `Punteado`.

- [ ] **Step 1: Escribir la prueba que falla**

Agregar a `tests/test_ui_smoke.py`:

```python
def test_controles_de_bordes(report_mini, tmp_path, monkeypatch):
    monkeypatch.setenv("APPDATA", str(tmp_path))
    root = tk.Tk()
    app = App(report_mini, tmp_path / "stock17-06.xlsx", root=root)
    try:
        assert app.v_cell_border.get() == "Ninguno"
        assert app.v_table_border.get() == "Ninguno"
        app.v_cell_border.set("Grueso")
        app.v_table_border.set("Punteado")
        app._sync_format()
        assert app.fmt.cell_border == "grueso"
        assert app.fmt.table_border == "punteado"
    finally:
        root.destroy()
```

- [ ] **Step 2: Correr la prueba para ver que falla**

Run: `python -m pytest tests/test_ui_smoke.py::test_controles_de_bordes -q`
Expected: falla con `AttributeError: 'App' object has no attribute 'v_cell_border'`.

- [ ] **Step 3: Implementar los controles**

En `stocky/ui.py`:

- Agregar `BORDERS = {"Ninguno": "ninguno", "Fino": "fino", "Medio": "medio", "Grueso": "grueso", "Punteado": "punteado"}`.
- Crear `self.v_cell_border` y `self.v_table_border` con `_label(BORDERS, ...)`.
- Agregar dos `ttk.Combobox` readonly en `variable_widgets` y dos filas en el panel: `Líneas de celdas` y `Contorno de la tabla`, después de `Color del encabezado`.
- Agregar ambas variables al tuple de `trace_add`.
- En `_sync_format`, asignar `self.fmt.cell_border = BORDERS.get(self.v_cell_border.get(), "ninguno")` y lo mismo para `table_border`.
- En `_load_vars_from_format`, setear ambas con `_label(BORDERS, ...)`.

- [ ] **Step 4: Correr la prueba y la suite**

Run: `python -m pytest tests/test_ui_smoke.py::test_controles_de_bordes -q`
Expected: 1 passed.

Run: `python -m pytest -q`
Expected: toda la suite pasa.

- [ ] **Step 5: Verificar el archivo real y documentar**

Run: `python -m pytest tests/test_parser.py::test_reporte_de_columnas_real -q`
Expected: 1 passed.

Construir un xlsx del área 08 del archivo de octubre con `Formato(cell_border="fino", table_border="grueso")`, exportarlo a PDF y comprobar que el PDF existe y que el xlsx conserva los bordes. Agregar a `docs/superpowers/plans/verificacion-2026-10-22.md` una sección “Bordes de la tabla” con el resultado.

- [ ] **Step 6: Commit**

```bash
git add stocky/ui.py tests/test_ui_smoke.py docs/superpowers/plans/verificacion-2026-10-22.md
git commit -m "feat: agrega controles de bordes al panel"
```
