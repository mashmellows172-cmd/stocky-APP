# Bordes de la tabla — Diseño

**Fecha:** 2026-10-22
**Estado:** aprobado por el usuario (dos selectores independientes)

## Intención

Agregar al panel de formato dos controles opcionales para marcar la tabla
impresa:

- **Líneas de celdas:** una grilla alrededor de cada celda del encabezado y de
  las filas de datos.
- **Contorno de la tabla:** un borde exterior que rodea el encabezado y todas
  las filas de datos.

Cada control ofrece **Ninguno, Fino, Medio, Grueso y Punteado** de forma
independiente. Los dos valores se guardan en el formato y en los presets, y se
aplican tanto a la vista previa como a la exportación y a la impresión.

## Estado actual

`stocky/export.py` solo dibuja una línea fina debajo del encabezado. Los
"Márgenes" que ya existen en el panel son los márgenes de la hoja (0,8 / 1,0 /
1,5 / 2,0 cm), no los bordes de las celdas. El cambio no toca esos márgenes de
página.

## Comportamiento

- `Formato.cell_border = "ninguno"` y `Formato.table_border = "ninguno"` por
  defecto: la salida actual no cambia.
- `cell_border` distinto de `ninguno` aplica izquierda, derecha, arriba y abajo
  a todas las celdas de la tabla, incluidas las filas de códigos alternativos.
- `table_border` distinto de `ninguno` aplica el estilo elegido al perímetro:
  arriba del encabezado, abajo de la última fila, izquierda de la primera
  columna y derecha de la última. El perímetro pisa el estilo de celda en esos
  lados.
- Si `cell_border` es `ninguno`, se conserva la línea fina debajo del encabezado
  que ya existe. Si hay borde de celda, esa línea usa el estilo elegido.
- El título, la línea de metadatos y los márgenes de hoja no llevan borde de
  tabla.
- El color de todas las líneas nuevas es el gris actual (`808080`), para
  mantener la estética y que se lea bien en impresión.

## Enfoques considerados

1. **Bordes por celda en el xlsx (elegido).** Se aplican estilos de openpyxl a
   las celdas y al perímetro. Es el mismo archivo que ya alimenta la vista
   previa, la exportación y la impresión, y permite controlar cada lado.
2. **Tabla de Excel con estilo.** Menos código, pero ata la apariencia a estilos
   de Excel, no distingue bien el perímetro del interior y puede cambiar al
   abrir el archivo en otra versión. Se descarta.
3. **Dibujar sobre el PDF.** Permitiría efectos que openpyxl no ofrece, pero
   habría que repetir la lógica en cada exportación y la vista previa podría no
   coincidir con el xlsx. Se descarta.

## Datos y UI

- `stocky/presets.py`: agregar `cell_border` y `table_border` a `Formato`, con
  valores permitidos y saneamiento de datos viejos o inválidos.
- `stocky/export.py`: mapear los nombres a estilos de openpyxl (`thin`,
  `medium`, `thick`, `dotted`) y aplicar los bordes después de escribir las
  filas, para cubrir también las filas de códigos alternativos.
- `stocky/ui.py`: dos combobox "Líneas de celdas" y "Contorno de la tabla";
  ambos disparan la actualización de la vista previa como el resto del formato.
- Los presets existentes siguen funcionando: al no traer los campos nuevos,
  toman `ninguno`.

## Pruebas

- Exportación: estilos de celda y perímetro en encabezado, datos y código
  alternativo; default sin cambios; cada tipo de línea mapea al estilo correcto.
- Presets: ida y vuelta de los dos campos y saneamiento de valores inválidos.
- UI: los controles existen, reflejan el formato y actualizan la vista previa.
- Suite completa y prueba real de octubre siguen verdes.

## Fuera de alcance

- Colores de línea, grosores personalizados, márgenes internos (padding) y
  bordes en el título o en los metadatos.
