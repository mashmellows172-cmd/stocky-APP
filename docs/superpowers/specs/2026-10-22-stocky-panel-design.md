# Stocky — panel de extracción y formato de stock diario

Fecha: 2026-10-22
Estado: borrador para aprobación

## 1. Intención

Cada día llega un archivo Excel (`stockDD-MM.xlsx`) que en realidad es un reporte de
texto del sistema SAI ("LISTADO DE EXISTENCIAS") volcado en la columna A. Hoy se abre,
se filtra a mano por área, se ordena y se imprime una tabla por área (8, 9, 29, 46, 84,
99 y a veces otras). Esas tablas se usan en sala para saber cuántas unidades hay de
cada producto y a qué precio, y así vender más rápido.

Se quiere reemplazar ese trabajo manual por: **arrastrar el archivo sobre un ícono,
ajustar el aspecto de la tabla si hace falta, ver cómo queda y generar/imprimir**.

Éxito = de descargar el archivo a tener las tablas impresas sin abrir Excel, sin
filtrar a mano y sin reescribir encabezados, con la misma información que hoy y mejor
legibilidad.

## 2. Objetivos

1. Un solo gesto: soltar el xlsx sobre el ícono y llegar a una vista previa real.
2. Ajustar el formato (fuente, tamaño, márgenes, densidad, columnas) y verlo al instante.
3. Presets: guardar un formato y reutilizarlo, incluso distinto por área.
4. Generar una carpeta con un archivo por área, listo para imprimir (xlsx + pdf).
5. Imprimir el área que se está viendo sin salir de la app.
6. Ser tolerante: si el reporte viene raro, decir qué línea falló y no generar basura.

## 3. No objetivos

- Editar stock, precios o altas/bajas. La app solo lee y presenta.
- Conectarse al SAI, a una base de datos o a la red.
- Leer otros reportes del SAI (solo "Listado de existencias").
- Multiusuario, nube, sincronización.
- Reemplazar la impresión del sistema: la salida es un documento listo para imprimir.

## 4. Anti-objetivos

- Que el usuario tenga que abrir Excel para completar el trabajo.
- Que la app toque o reescriba el archivo original descargado.
- Que se imprima sin haber visto antes cómo queda (sin vista previa).
- Que un formato distinto del esperado produzca tablas silenciosamente incompletas.

## 5. Restricciones y entorno

- Windows 10/11, un solo equipo, un solo usuario.
- Python 3.14.5 ya instalado; `openpyxl` y `pandas` disponibles; `tkinter` 8.6 disponible.
- Microsoft Excel 16 instalado y con COM funcionando (verificado).
- `pypdfium2` + `Pillow` instalables para renderizar el PDF (verificado por `pip --dry-run`);
  `pywin32` para hablar con Excel, y `pytest` para pruebas.
- Sin acceso a internet en tiempo de ejecución.
- Impresora: papel **carta (Letter)**. Orientación vertical por defecto y horizontal
  disponible como ajuste.
- El reporte real de referencia: 37.779 filas, 26.527 ítems, 69 áreas, 3.192 códigos
  alternativos.

## 6. Los datos (lo que el parser debe entender)

Cada línea del reporte es texto con campos separados por `|` (13 separadores, sin `|`
dentro de las descripciones — verificado sobre las 37.779 filas).

Línea de ítem:

```
 |  1   |7881 MEDIANA                |26726760  |00200026726766 | AMERICAN TOURISTER MALETA FRANKFURT. |(UNI)   -1|(UNI)   -1|       -1 |       -1 |     59990 |     -49994|
```

| Campo | Ejemplo | Nota |
|---|---|---|
| AREA | `1` | número de área (no trae nombre) |
| NIVEL DE AGREGACION | `7881 MEDIANA` | rubro: código + nombre |
| PLU | `26726760` | 8 dígitos, conservar como texto |
| CODIGO DE BARRA | `00200026726766` | 14 dígitos, ceros a la izquierda significativos |
| DESCRIPCION | `AMERICAN TOURISTER MALETA FRANKFURT.` | máximo 38 caracteres, ya viene truncada |
| UN. EMP | `(UNI)   -1` | unidad de empaque + cantidad |
| UN. DESPA | `(UNI)   -1` | ídem |
| EXISTENCIA | `-1` | entero, puede ser negativo |
| DISPONIBLE | `-1` | entero; **es el que manda** |
| PRECIO VENTA | `59990` | entero en pesos, sin decimales |
| COSTO PROMEDIO | `-49994` | entero, puede ser negativo |

Particularidades que hoy rompen el trabajo a mano:

1. **Segundo código de barra.** Un ítem con más de un código ocupa dos líneas; la
   segunda trae solo el código alternativo (3.192 casos en el archivo real). Debe
   quedar asociado al ítem anterior, no como producto aparte.
2. **Encabezados por página.** Cada página repite 4 líneas de encabezado y 2 reglas de
   guiones. Se descartan.
3. **Metadatos** en el encabezado: `FECHA : 17/06/2026`, `HORA : 10:06:36`,
   `DEPENDENCIA : 10 MAIPU`, `FECHA EMISION`. Se usan para el encabezado de la tabla.
4. **Acentos rotos.** El export pierde las vocales acentuadas (`�rea`). No se recuperan;
   el texto se normaliza a mayúsculas ASCII para imprimir.
5. **El nombre del área no viene en el archivo.** Sale de un mapa fijo (§7).
6. **Stock negativo y cero.** 1.759 ítems con disponible < 0 y 188 en 0. No se imprimen.
7. **Filas que se desbordan a la celda B.** La fila 37.067 del archivo real viene
   partida entre las columnas A y B. El parser concatena **todas las celdas pobladas**
   de la fila antes de interpretarla; de lo contrario se pierde ese ítem.
8. **Reportes vacíos o de otro formato.** `stock19-10.xlsx` tiene 44.032 filas de puro
   encabezado y cero ítems. Si el parser no encuentra al menos un ítem válido, la app
   no abre el panel: avisa "el reporte no trae ítems" con la ruta del archivo. Un
   export de columnas separadas (AREA, RUBRO, PLU... en celdas distintas) queda fuera
   de alcance hasta tener un ejemplo real con datos; se detecta y se informa como
   formato no soportado, no se procesa a medias.

## 7. Mapa de áreas (letrero de la dependencia 10 MAIPU)

| N° | Nombre | N° | Nombre |
|---|---|---|---|
| 8 | AUDIO | 78 | PYMES |
| 9 | TELEFONIA | 79 | REFRIGERACION |
| 10 | ELECTRODOMESTICOS | 80 | HORNO |
| 29 | ELECTRONICA MENOR | 81 | LAVADO Y SECADO |
| 46 | COMPUTACION | 83 | ASPIRADO Y LIMPIEZA |
| 73 | ACCESORIOS COMPUTACION | 84 | COCINA |
| 74 | ENTRETENCION | 85 | OTROS LINEA BLANCA |
| 77 | CUIDADO PERSONAL | 86 | PERIFERICOS |
| | | 95 | TEMPORADA |
| | | 99 | TELEVISION |

El letrero repite "80 HORNO"; en los datos existe una sola área 80. Las áreas que no
estén en el mapa se generan igual, con el número como nombre (`37.xlsx`).

## 8. Enfoques considerados

| Enfoque | Qué era | Veredicto |
|---|---|---|
| A. Script + ícono para arrastrar | Sin ventana; salida directa | Base elegida: es el gesto diario |
| B. Ventana mínima | Elegir áreas y generar | Se descartó sola |
| C. Un libro con una hoja por área | Un archivo único | Descartado: imprimir exige elegir hoja |

**Elegido: A + B combinados.** Se arrastra el archivo y se abre un panel (maqueta A
aprobada): barra lateral de ajustes, vista previa de la página real a la derecha,
barra inferior con "Generar" e "Imprimir". El motor queda separado de la interfaz.

Sub-decisiones y por qué:

- **La vista previa es el PDF real.** El xlsx se construye, Excel lo exporta a PDF en
  segundo plano y cada página se muestra como imagen. Lo que se ve es exactamente lo
  que sale por la impresora, sin sorpresas de configuración de página.
- **Tkinter, no PySide.** Ya viene con Python, no suma 150 MB y alcanza de sobra para
  un panel con imágenes. El riesgo de compatibilidad con Python 3.14 desaparece.
- **La salida incluye xlsx y pdf.** El pdf es para imprimir; el xlsx queda por si hay
  que revisar o reimprimir con otro formato.

## 9. Cómo funciona

```
[arrastrar xlsx sobre el ícono]
        │
        ▼
parser.py  →  StockReport (meta + ítems + anomalías)
        │
        ▼
ui.py  →  panel con ajustes, lista de áreas y vista previa
        │            │
        │            └─ preview.py: export → Excel COM → PDF → imágenes (debounce 400 ms)
        ▼
[Generar]  →  export.py (xlsx) + pdf.py (PDF) por cada área marcada
[Imprimir] →  envía el PDF del área visible a la impresora predeterminada,
              o lo abre para elegir impresora
```

Salida por defecto, junto al archivo original:

```
Stock 2026-06-17/
  08 - AUDIO.xlsx       08 - AUDIO.pdf
  09 - TELEFONIA.xlsx   09 - TELEFONIA.pdf
  99 - TELEVISION.xlsx  99 - TELEVISION.pdf
```

Si la carpeta existe, se reutiliza y se sobrescriben los archivos de las áreas marcadas.

## 10. Contenido de cada tabla

Columnas, en este orden: **PLU · CÓDIGO · DESCRIPCIÓN · DISPONIBLE · PRECIO**. Opcional
(un check): **RUBRO** entre código y descripción.

Encabezado de la tabla (arriba de cada página, repetido):

```
ÁREA 09 · TELEFONIA                    MAIPU · 17/06/2026 · 10:06
```

Pie de página, repetido en cada página: `Página X de Y` y
`N ítems · U unidades` (cantidad de renglones y suma de DISPONIBLE de toda el área).

Reglas de negocio fijas:

- Solo `disponible >= 1`.
- Orden: descripción ascendente (sin distinguir mayúsculas ni acentos), desempate por PLU.
- Un renglón por ítem; los códigos alternativos van en una segunda línea chica bajo el
  código principal, sin sumar unidades.
- Precio con separador de miles (`1.699.990`) salvo que se desactive en el panel.
- El código de barra conserva los ceros a la izquierda y se imprime en texto.

Página: **carta (Letter)**, vertical por defecto. La tabla siempre se ajusta a una hoja
de ancho (`FitToPagesWide = 1`); si el ajuste elegido no entra, la app reduce la escala
antes que partir columnas en dos páginas.

## 11. Ajustes del panel (personalización)

| Ajuste | Valores | Por defecto |
|---|---|---|
| Fuente | lista de fuentes instaladas + sugeridas (Segoe UI, Calibri, Arial, Tahoma, Verdana, Consolas) | Segoe UI |
| Tamaño del cuerpo | 8–16 pt | 10 |
| Tamaño del título | 1–2 pasos sobre el cuerpo, o fijo | cuerpo + 4 |
| Densidad de fila | Compacta / Normal / Amplia | Normal |
| Filas alternadas | sí / no | no |
| Color de encabezado | sin relleno / gris claro / verde stock | gris claro |
| Márgenes | 0,8 / 1,0 / 1,5 / 2,0 cm | 1,5 cm |
| Orientación | vertical / horizontal | vertical |
| Ancho de descripción | normal / ancha | normal |
| Mostrar rubro | sí / no | no |
| Separador de miles | sí / no | sí |
| Encabezado extra | texto libre | vacío |

Los ajustes son **globales**: lo que se mueve en el panel se aplica a todas las áreas a
la vez. Un preset guarda esa combinación con un nombre ("Sala", "Oficina", "Letra
grande") y se aplica desde una lista. Botones: **Guardar como preset**, **Presets ▾**,
**Restablecer**.

Persistencia: `%APPDATA%\Stocky\config.json` (biblioteca de presets con nombre, últimos
ajustes, últimas áreas marcadas y última carpeta usada). No se escribe nada en la
carpeta del proyecto.

## 12. Manejo de errores

| Situación | Qué hace la app |
|---|---|
| El archivo no es un "Listado de existencias" | Aviso, no abre el panel |
| Reporte sin ítems (solo encabezados) | Aviso "el reporte no trae ítems", no abre el panel |
| Export de columnas separadas (formato no soportado) | Aviso con la ruta; no procesa a medias |
| Líneas con formato desconocido | Las cuenta y muestra en el pie ("3 líneas no reconocidas"), el resto se procesa |
| PLU repetido con mismos datos y otro código | Se fusiona: el segundo código pasa a alternativo |
| PLU repetido con datos distintos | Conserva el primero, informa "1 duplicado conflictivo" |
| Área sin ítems con stock | No genera archivo; lo indica en la lista |
| Excel COM no responde | Genera los xlsx igual y avisa que el pdf/vista previa no está disponible |
| Archivo abierto o carpeta protegida | Mensaje con la ruta y el motivo |

## 13. Arquitectura

```
stocky-APP\
  Stocky.lnk              acceso directo que recibe el archivo arrastrado
  abrir consola.cmd       lanzador con consola, para diagnosticar
  stocky\
    app.py                entrada: argumento arrastrado, arranque de la UI
    models.py             StockRow y StockReport
    parser.py             texto del reporte → StockReport
    areas.py              mapa número → nombre
    export.py             StockReport + Formato → xlsx (openpyxl)
    pdf.py                xlsx → PDF (Excel COM), envío a impresora
    preview.py            PDF → imágenes de página (pypdfium2) + caché
    presets.py            Formato y config.json, biblioteca de presets
    ui_state.py           selección de áreas, textos de estado, fuente válida
    theme.py              colores, fuentes, estilos ttk
    ui.py                 ventana, panel lateral, vista previa
  tests\
    fixtures\             recorte real del reporte (encabezado, ítem, continuación)
    test_parser.py
    test_areas.py
    test_export.py
    test_pdf.py
    test_preview.py
    test_presets.py
    test_ui_state.py
    test_ui_smoke.py
    test_app.py
  docs\superpowers\specs\ este documento
```

`Formato` es un dataclass con los valores de §11; `StockReport` es metadatos + lista de
`StockRow` (area, rubro, plu, barcode, alt_barcodes, desc, un_emp, un_despa, existencia,
disponible, precio, costo) + lista de anomalías. El parser no depende de la UI y se
prueba solo.

## 14. Pruebas

- Parser: recorte real de 40 líneas (encabezado de página, ítem normal, ítem con dos
  códigos, ítem negativo, fila desbordada a la celda B, área distinta) → mismos valores
  que el archivo original.
- Parser sobre el archivo completo: 26.527 ítems, 69 áreas, 3.192 códigos alternativos,
  y el ítem de la fila desbordada presente.
- Reporte vacío (`stock19-10.xlsx`) → error "no trae ítems", sin panel.
- Deduplicación: PLU con mismo stock y otro código se fusiona; PLU con datos distintos
  se informa como conflictivo.
- Export: leer el xlsx generado y verificar orden, filtro, ceros a la izquierda,
  formato de precio, ajuste a una hoja de ancho y pie con ítems/unidades.
- Presets: guardar, leer, aplicar y sobrevivir a un JSON corrupto.
- Manual: arrastrar el archivo real sobre el ícono, ver la vista previa, generar e imprimir.

## 15. Decisiones tomadas por mí (revisar)

1. La vista previa muestra un área a la vez (pestañas por área); renderizar las 6 juntas
   sería lento y no agrega nada.
2. Dos botones de impresión: "Imprimir" (directo a la predeterminada) y "Abrir PDF".
3. Filas alternadas apagadas por defecto: menos tinta y el papel actual no las tiene.
4. Salida junto al archivo original, en subcarpeta con la fecha del reporte.
5. Un formato global + biblioteca de presets con nombre, guardados en `%APPDATA%`.
   Se descarta el formato distinto por área (era de la maqueta C, no la elegida).
6. Tkinter y PDF real como fuente de verdad de la vista previa.
7. Defaults: carta vertical, Segoe UI 10, márgenes 1,5 cm, precio con puntos de miles.
   Vertical es lo óptimo para 5 columnas con descripción de 38 caracteres; horizontal
   queda disponible para quien quiera descripciones más anchas.
8. El filtro de stock (`disponible >= 1`) es fijo, no configurable.
9. Reportes de otro formato (columnas separadas) se rechazan con aviso hasta tener un
   ejemplo real con datos; no se adivina el formato.

## 16. Lo que queda al constructor

- Estructura interna de cada módulo y nombres de funciones.
- Cómo se dibuja el zoom y el desplazamiento de la vista previa.
- Estrategia de caché de imágenes y de archivos temporales.
- Cómo se genera el `.ico` y cómo se crea el acceso directo.
- Detalles de estilo fino del panel (espaciados, iconos).
