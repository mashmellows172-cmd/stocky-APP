# Stocky portable — Diseño

**Fecha:** 2026-10-23
**Estado:** aprobado por el usuario: quiere ejecutarlo en una PC externa sin depender de Python.

## Intención

Entregar `Stocky.exe`, un ejecutable único para Windows 64 bits que incluya el
intérprete de Python y todas las dependencias. Se copia a otra PC, se abre con
doble clic o arrastrando un reporte `.xlsx`, y funciona sin instalar Python.

## Objetivos

- Un solo archivo `Stocky.exe`, sin consola visible.
- Incluir Tkinter, openpyxl, Pillow, pypdfium2 y pywin32.
- Mantener el comportamiento actual: panel, presets, vista previa, guardado,
  exportación e impresión.
- Poder copiarlo por OneDrive o USB y ejecutarlo desde cualquier carpeta.
- Dejar un build reproducible en `herramientas/construir_exe.py`.

## No objetivos

- No incluye Excel: para PDF e impresión la PC externa necesita Microsoft Excel
  instalado. Guardar el `.xlsx` sí funciona sin Excel.
- No es un instalador, no crea accesos directos y no se firma digitalmente.
- No soporta Windows de 32 bits.
- No modifica la lógica de la aplicación ni los márgenes, bordes o presets.

## Enfoques considerados

1. **PyInstaller onefile (elegido).** Genera un único `.exe` que se extrae a una
   carpeta temporal al arrancar. Es lo más simple para copiar a otra PC; el
   arranque es algo más lento y el archivo pesa más.
2. **PyInstaller onedir.** Genera una carpeta con el `.exe` y las DLL. Arranca
   más rápido y es más fácil de diagnosticar, pero son muchos archivos y se
   puede romper al copiar solo el `.exe`.
3. **Python embebido + lanzador.** Una carpeta con Python portátil y un `.bat` o
   `.vbs`. Evita PyInstaller, pero deja muchos archivos y un lanzador que
   depende de rutas relativas.

Se elige el enfoque 1 porque el requisito principal es que sea un archivo fácil
de llevar a la PC externa.

## Decisiones de implementación

- Build con Python 3.14.5 y PyInstaller 6.22.3, en Windows 64 bits.
- Comando base:
  `python -m PyInstaller --noconfirm --clean --onefile --windowed --name Stocky --icon assets/stocky.ico stocky/app.py`
- Si PyInstaller no detecta alguna DLL de `pypdfium2` o `pywin32`, se agregan
  `--collect-all pypdfium2` y los hidden imports necesarios; el build se valida
  ejecutando el `.exe` con un reporte real, no solo compilando.
- El ejecutable queda en `dist/Stocky.exe`, que está ignorado por git; el
  usuario lo copia desde ahí. El script de build puede generar además un
  `dist/LEEME Stocky.txt` con instrucciones.
- Se conserva `Stocky.lnk` como acceso para este PC; en la PC externa se usa
  directamente `Stocky.exe`.

## Verificación

- `python -m pytest -q` sigue pasando.
- `dist/Stocky.exe` abre el panel con un `.xlsx` real y muestra el título del
  reporte.
- Un smoke test frozen comprueba que el ejecutable empaquetado puede leer el
  Excel, escribir un xlsx y generar el PDF con Excel.
- Se verifica el tamaño y que no haya dependencias externas de Python.

## Riesgos

- SmartScreen puede advertir por un ejecutable nuevo sin firma; se ejecuta con
  “Más información → Ejecutar de todas formas”.
- Sin Excel en la PC externa, la vista previa/PDF fallará con un mensaje claro;
  el xlsx se puede generar igual.
- El onefile se extrae en `%TEMP%`; antivirus corporativos pueden demorar el
  primer arranque.
