# Verificación final — 2026-10-22

Corrida hecha sobre `C:\Users\victus\Downloads\stock17-06.xlsx` (reporte real) y
`stock19-10.xlsx` (reporte vacío).

## Suite completa

`python -m pytest -q` → **45 passed**, sin pruebas saltadas. Corren la prueba `slow`
contra el archivo real y las dos de Excel/PDF.

## Generación de las seis áreas

Salida en `C:\Users\victus\Downloads\Stock 2026-06-17\`: **12 archivos**
(6 `.xlsx` + 6 `.pdf`).

| Área | Ítems | Unidades | Páginas |
|---|---|---|---|
| 08 - AUDIO | 31 | 99 | 1 |
| 09 - TELEFONIA | 103 | 755 | 3 |
| 29 - ELECTRONICA MENOR | 125 | 1.067 | 4 |
| 46 - COMPUTACION | 39 | 133 | 1 |
| 84 - COCINA | 16 | 29 | 1 |
| 99 - TELEVISION | 56 | 120 | 2 |

## Cifras contra el parser

Área 9: el xlsx generado tiene **103 filas de datos** y **755 unidades**, igual que
`rows_for(report, 9)`. El pie impreso dice `103 ítems · 755 unidades` y
`Página &P de &N`.

## Columnas completas

Se extrajo el texto de **todas** las páginas de las seis áreas: en ninguna falta
`PLU`, `CÓDIGO`, `DESCRIPCIÓN`, `DISPONIBLE` ni `PRECIO`.

## Peor caso de formato

Área 9 con Segoe UI 16, descripción ancha y márgenes de 2,0 cm: **2 páginas**, las
cinco columnas completas y sin texto pisado. Al agrandar la letra se ensanchan las
columnas y Excel reduce la escala de impresión antes que partir la tabla
(`FitToPagesWide = 1`).

## Caminos de error

- `stock19-10.xlsx` → `Formato no soportado: no parece un Listado de existencias`,
  salida 1, sin ventana ni carpeta nueva. El plan esperaba `El reporte no trae
  ítems`; ese texto queda para archivos que sí parecen listado pero no traen filas.
  El archivo vacío no trae ni una línea con `|`, por eso se detecta antes como
  formato desconocido.
- `config.json` con basura → la app arranca con defaults (`presets {}`, fuente 10).

## Arrastre

`pythonw stocky\app.py "…\stock17-06.xlsx"` abre la ventana **Stocky - 17/06/2026**
con las seis áreas de siempre marcadas y la vista previa del área 08.

## Diferencia encontrada y corregida

En el peor caso la letra de 16 pt se pisaba porque el ancho de columna de Excel se
mide en caracteres del tamaño base. Ahora los anchos se escalan con el tamaño de
letra (`font_size / 11`).

## Revisión final: las cinco correcciones aplicadas

Después de la revisión independiente se corrigieron los cinco hallazgos de mayor
impacto:

1. **Imprimir.** El visor de PDF de Edge no registra el verbo `print`; por eso
   `os.startfile(pdf, "print")` fallaba con WinError 1155. Ahora el panel regenera
   el `.xlsx` del área y lo imprime con Excel (`Workbook.PrintOut`) en un hilo
   aparte. La impresión usa el mismo formato que la vista previa.
2. **Errores del panel.** Los `lambda` de error usaban la variable del `except`,
   que Python borra al salir del bloque, y mostraban `NameError`. El texto se
   captura antes de programar el aviso.
3. **Archivo equivocado.** Un `.xlsx` corrupto o un `.xls` viejo ya no cierran
   `pythonw`: se muestra `No se pudo leer el reporte …` y se sale con código 1.
4. **Líneas con pipes inválidas.** Las líneas que tienen pipes pero no la forma
   esperada se cuentan en `unknown_lines` en vez de descartarse en silencio.
5. **Alto de fila.** Ahora escala con la letra (compacta 1,2 / normal 1,5 /
   amplia 1,8) y el código alternativo usa 1,2 × su propia fuente.

Evidencia: suite **45 passed**; `PrintOut` verificado sobre un libro real abierto
con Excel; el peor caso (16 pt + compacta + códigos alternativos) se renderizó sin
texto recortado: `evidencia-16pt-compacta.png`.
