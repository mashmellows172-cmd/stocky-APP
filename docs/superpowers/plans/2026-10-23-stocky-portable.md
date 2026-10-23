# Stocky portable Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Producir un `Stocky.exe` único de 64 bits que funcione en una PC externa sin Python instalado.

**Architecture:** PyInstaller empaqueta `stocky/app.py` en modo onefile/windowed con Tkinter, openpyxl, Pillow, pypdfium2 y pywin32. Un script de build reproducible genera el ejecutable y un LEEME; la verificación se hace ejecutando el binario con el reporte real, no solo compilando.

**Tech Stack:** Python 3.14.5, PyInstaller 6.22.3, Tkinter, openpyxl, Pillow, pypdfium2, pywin32, Excel COM.

**Spec:** `docs/superpowers/specs/2026-10-23-stocky-portable-design.md`

## Global Constraints

- Windows 64 bits; el ejecutable es onefile y sin consola.
- Incluir Python y dependencias; no instalar Python en la PC externa.
- Excel sigue siendo necesario solo para PDF e impresión.
- No modificar la lógica de la aplicación.
- Salida: `dist/Stocky.exe`; `dist/` está ignorado por git.
- Copiar el exe por OneDrive o USB y probarlo desde una carpeta distinta.

## Review Focus

1. **Onefile y rutas temporales:** el exe debe abrir desde cualquier carpeta y no depender de `Stocky.lnk`.
2. **Dependencias ocultas:** pywin32 y pypdfium2 deben estar dentro del binario; el PDF debe funcionar.
3. **Rutas con espacios/acentos:** el reporte y la carpeta de salida pueden tener ambos.
4. **Arrastrar y soltar:** soltar un `.xlsx` sobre `Stocky.exe` debe abrir el panel.
5. **Sin Excel:** el panel debe seguir abriendo y guardando xlsx; PDF/impresión deben mostrar error claro.

---

### Task 1: Build reproducible del ejecutable

**Files:**
- Create: `herramientas/construir_exe.py`
- Create: `docs/superpowers/specs/2026-10-23-stocky-portable-design.md` (ya escrito)
- Output: `dist/Stocky.exe`

**Interfaces:**
- Consumes: `stocky/app.py` como entry point.
- Produces: `construir_exe()` que devuelve la ruta del ejecutable generado.

- [ ] **Step 1: Instalar PyInstaller**

Run: `python -m pip install pyinstaller==6.22.3`
Expected: instala PyInstaller y sus dependencias.

- [ ] **Step 2: Escribir el script de build**

`herramientas/construir_exe.py` debe:

- Calcular la raíz del proyecto con `Path(__file__).resolve().parents[1]`.
- Ejecutar `sys.executable -m PyInstaller` con: `--noconfirm`, `--clean`, `--onefile`, `--windowed`, `--name Stocky`, `--icon assets/stocky.ico`, `stocky/app.py`.
- Usar `--collect-all pypdfium2` si el build inicial deja fuera la DLL de renderizado.
- Verificar que `dist/Stocky.exe` exista y devolver esa ruta.
- Escribir `dist/LEEME Stocky.txt` con instrucciones de uso y la aclaración de Excel.

- [ ] **Step 3: Ejecutar el build**

Run: `python herramientas/construir_exe.py`
Expected: termina con código 0 y crea `dist/Stocky.exe`.

- [ ] **Step 4: Si falla el análisis de imports**

Revisar `build/Stocky/warn-Stocky.txt`; agregar los hidden imports necesarios (`win32timezone`, `pythoncom`, `pywintypes`) o `--collect-all pypdfium2`; repetir el build.

- [ ] **Step 5: Commit**

```bash
git add herramientas/construir_exe.py docs/superpowers/specs/2026-10-23-stocky-portable-design.md docs/superpowers/plans/2026-10-23-stocky-portable.md
git commit -m "build: agrega empaquetado portable de Stocky"
```

### Task 2: Verificar el binario con el reporte real

**Files:**
- Test: ejecución manual de `dist/Stocky.exe`
- Test: smoke frozen temporal (sin commit)

**Interfaces:**
- Consumes: `dist/Stocky.exe`.
- Produces: evidencia de que el binario abre, lee el reporte y genera PDF sin Python externo.

- [ ] **Step 1: Abrir el exe con el reporte**

Run: `Start-Process .\dist\Stocky.exe -ArgumentList '"C:\Users\victus\Downloads\stock19-10.xlsx"'`
Expected: aparece la ventana de Stocky con el área 08 disponible; el proceso vive.

- [ ] **Step 2: Verificar el título y el panel**

Comprobar que el título contiene `Stocky` y la fecha del reporte, y que el panel muestra los controles de bordes. Cerrar la ventana.

- [ ] **Step 3: Smoke frozen de PDF**

Crear un script temporal que importe `read_report`, `write_area_xlsx`, `export_pdf` y `render_pages`; empaquetarlo con PyInstaller en modo consola y ejecutarlo con el reporte real. Expected: escribe un xlsx, genera un PDF de 1 página y termina con código 0.

- [ ] **Step 4: Probar desde otra carpeta**

Copiar `dist/Stocky.exe` a `%TEMP%\StockyPortable` y abrirlo con el reporte real. Expected: funciona sin `Stocky.lnk` ni la carpeta del proyecto.

- [ ] **Step 5: Documentar y commit**

Agregar a `docs/superpowers/plans/verificacion-2026-10-22.md` una sección “Stocky portable” con tamaño del exe, resultado del PDF y la aclaración de Excel. Commit solo de documentación y script; no commitear `dist/`.
