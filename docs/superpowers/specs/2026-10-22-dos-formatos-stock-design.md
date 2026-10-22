# Soporte de los dos formatos de reporte de stock

## Intención

Stocky debe abrir los dos formatos de “Listado de existencias” que se usan en
la tienda, detectando el formato solo al arrastrar el archivo:

- **Formato histórico:** cada ítem es una línea de texto con 13 campos
  separados por `|`.
- **Formato habitual (columnas):** una tabla con ÁREA, NIVEL DE AGREGACIÓN,
  PLU, CÓDIGO, DESCRIPCIÓN, DISPONIBLE y PRECIO, precedida por encabezados
  repetidos, espacios y filas de ruido.

El formato de columnas será el principal; el de pipes debe seguir funcionando
sin cambios.

## Objetivos

- Detectar automáticamente cuál de los dos formatos trae el archivo.
- Leer el formato de columnas sin depender de que la tabla empiece en una fila
  fija: saltar encabezados repetidos, filas vacías y filas auxiliares.
- Producir el mismo `StockReport`/`StockRow` que hoy, para que vista previa,
  filtros, presets, exportación e impresión no cambien.
- Mantener intacto el comportamiento del formato de pipes.
- Extraer fecha, hora y dependencia también del formato de columnas.
- Reutilizar la deduplicación de PLU y el conteo de conflictos existentes.
- Dejar pruebas de los dos formatos y una prueba contra el archivo real de
  octubre.

## No objetivos

- Agregar un selector de formato en la interfaz: la detección es automática.
- Soportar otros formatos de reporte además de estos dos.
- Recuperar acentos que ya vienen perdidos como `�` en el archivo de origen.
- Cambiar la exportación, la vista previa, los presets o la impresión.
- Reconstruir las columnas que el formato de columnas no trae (unidades en
  empresa, unidades en despacho, existencia y costo).

## Contexto encontrado

El archivo `stock19-10.xlsx` no está vacío ni corrupto:

- 44.032 filas en total.
- Unas 6.320 filas iniciales de encabezados repetidos/espacios; la tabla real
  empieza en la fila 6.321.
- 30.500 filas de ítems y 30.424 PLU distintos (76 PLU repetidos).
- 66 áreas con ítems. Después de deduplicar y filtrar `disponible >= 1`, las
  seis de siempre quedan así:
  área 08: 39 ítems / 102 unidades; área 09: 107 / 615; área 29: 133 / 1.364;
  área 46: 36 / 111; área 84: 17 / 29; área 99: 74 / 212.
- Columnas: B = área, C = nivel de agregación (rubro), D = PLU, E = código de
  barra, F = descripción, G = disponible, H = precio.
- No trae unidades en empresa, despacho, existencia ni costo.

Hoy `read_report` junta las celdas de cada fila y exige encontrar `|`; por eso
rechaza el formato de columnas con “Formato no soportado”.

## Enfoques considerados

1. **Dos parsers con detección automática (elegido).** Se conserva el parser
   de pipes y se agrega un parser de tabla. `read_report` detecta el formato y
   deriva al parser correspondiente. Aísla el formato nuevo, mantiene el viejo
   sin tocar y permite probar cada uno por separado.
2. **Convertir la tabla a líneas con pipes y reutilizar el parser actual.**
   Menos código nuevo, pero obliga a fabricar un formato intermedio, maneja los
   metadatos por fuera y hace más difícil detectar datos raros.
3. **Un parser único que decida fila por fila.** Menos archivos, pero mezcla
   los dos formatos en la misma función y complica el diagnóstico de errores.

Se eligió el enfoque 1.

## Diseño

### Lectura y detección

`read_report` lee la hoja una sola vez como filas de celdas y reúne:

- **Señal de tabla:** alguna fila con celdas separadas que, ya recortadas,
  incluyan `PLU` y `DESCRIPCION PRODUCTO`.
- **Señal de pipes:** alguna fila cuyo texto unido contenga `|`.

La detección prioriza la señal de tabla. Así, un `|` perdido dentro de una
descripción no hace que un archivo de columnas se interprete como el formato
viejo. Los archivos de pipes no tienen esas dos celdas de encabezado separadas.

- Señal de tabla → `parse_table_rows(rows)`.
- Si no, señal de pipes → `parse_rows(lines)` (parser actual).
- Si no hay ninguna → `ParseError("Formato no soportado: no parece un Listado
  de existencias")`, igual que hoy.

### Parser de columnas

Recorre todas las filas y acepta como ítem solo las que cumplen:

- área numérica en B;
- PLU de 5 a 8 dígitos en D;
- código no vacío en E;
- descripción no vacía en F;
- disponible y precio numéricos, o vacíos (se interpretan como 0).

Las filas de encabezado, espacios, códigos sueltos y filas auxiliares se
ignoran. Los campos que el formato no trae (`un_emp`, `un_despa`, `existencia`,
`costo`) quedan en 0; la app solo usa `disponible` y `precio` para mostrar,
filtrar y exportar.

`rubro` se toma de C. La fecha, hora y dependencia se extraen de las primeras
filas con el mismo `_read_metadata` que usa el formato de pipes.

Al final se aplica la misma `_dedupe` y, si no quedó ninguna fila, se lanza
`ParseError("El reporte no trae ítems")`.

### Lo que queda a criterio de la implementación

- Nombres exactos de las funciones auxiliares y si la detección vive en
  `read_report` o en un helper.
- Cómo se recorre la hoja con bajo consumo de memoria.
- Detalle del fixture de pruebas y qué filas de ruido incluye.
- Si conviene contar como `unknown_lines` alguna fila con área y PLU pero sin
  descripción. El archivo real no debe inflar ese número por encabezados.

## Pruebas

- Fixture nuevo de formato columnas con metadatos, encabezados repetidos,
  filas vacías, un PLU repetido con otro código (debe fusionarse como código
  alternativo), un ítem con disponible 0 y un ítem normal.
- El test del formato pipes existente debe seguir pasando sin cambios.
- Detección: el mismo `read_report` acepta ambos fixtures.
- Prueba lenta contra `stock19-10.xlsx`: 30.424 filas únicas, 66 áreas y los
  conteos de las seis áreas de siempre.
- Prueba de archivo desconocido: sigue dando el error visible actual.
