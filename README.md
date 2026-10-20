# Control de órdenes: procesamiento del reporte SAP

## Aplicación «Seguimiento Ordenes.exe» (uso diario)
Para abrirla, haz doble clic en el `.exe`. No necesita Python. Tarda unos segundos en arrancar.

1. **📥 Cargar CSV de SAP…**: elige el CSV (o el xlsx) descargado de SAP. La aplicación hace cuatro cosas:
   - limpia la data;
   - genera `Control de ordenes DD_MM.xlsx` en esta carpeta, con el mismo libro de siempre más la hoja **Seguimiento**;
   - guarda una copia del CSV en `Cargas\` para conservar el historial;
   - actualiza el seguimiento:
     - solo entran las **OC sin factura que están En proceso**. Las Cerradas o Canceladas sin factura ya no se van a facturar y no se tratan;
     - si en el nuevo reporte la OC ya trae factura, se cierra como **Facturada** y deja anotado el N° de factura;
     - si la OC pasa a Cerrada o Cancelada en SAP sin factura, se cierra como **OC cerrada sin factura** u **OC cancelada sin factura**, conservando sus comentarios;
     - si la OC ya no aparece en el reporte, se cierra como **Fuera del reporte**;
     - si la factura se anula, la OC vuelve a seguimiento;
     - los cambios de Estado de la OC quedan registrados.
2. **Seguimiento**: en la lista, elige una OC y completa:
   - estado de seguimiento, responsable y fecha compromiso, y pulsa **Guardar cambios**;
   - **comentarios**, con **Agregar comentario** o Ctrl+Enter.

   Todo queda en el **Historial** con fecha y autor, y no se borra.
3. **Pestaña «Facturas»**: muestra todas las facturas del último reporte, una fila por factura.
   - Por defecto filtra **Pendiente de pago**. También puedes ver Pagado, Todos o (sin estado).
   - Otros filtros: comprador, estado de la OC, moneda, búsqueda, solo vencidas sin pagar y solo con comentarios.
   - Columnas: importes, emisión, vencimiento, **días para vencer** (negativo = vencida, en rojo) y fecha de pago. Las pagadas se ven en verde.
   - El resumen muestra el **importe pendiente por moneda** de lo que estás viendo.
   - Puedes comentar la OC de la factura seleccionada: el historial es único por OC.
   - En cada carga se registra en el historial de la OC cuando una factura **se paga** (con su fecha de pago), cuando aparece una **factura nueva** y cuando una factura **desaparece** del reporte (posible anulación).
4. **🔄 Actualizar Excel con comentarios**: vuelve a generar el control del último reporte con los comentarios del momento. Si el Excel está abierto, se guarda una copia con la hora en el nombre.

Filtros disponibles: texto, comprador, estado OC, situación, estado de seguimiento, solo sin comentarios y solo compromisos vencidos.

Colores de la lista:
- amarillo: OC sin comentarios;
- rojo: compromiso vencido;
- verde: OC facturada.

Si eliges el estado **"Cerrar sin factura"**, la OC deja de estar activa. Úsalo cuando ya sabes que una OC en proceso no se va a facturar, aunque todavía no esté cerrada en SAP.

Archivos de la aplicación (deben quedarse junto al `.exe`):
- `seguimiento.db`: comentarios e historial. **No lo borres.**
- `config_seguimiento.json`: tu nombre y las últimas rutas.
- `Cargas\`: copias de los reportes cargados.

> La base está pensada para **una persona a la vez**. OneDrive sincroniza el archivo, pero si dos personas comentan al mismo tiempo desde equipos distintos, puede generarse un conflicto de versiones.

**Carga automática (opcional)**: con el Programador de tareas de Windows puedes ejecutar
`"Seguimiento Ordenes.exe" --cargar "C:\...\Control de ordenes.csv"`. El resultado queda en `ultima_carga.log`.

**Recompilar el .exe** (solo si cambias el código):
`python -m PyInstaller --onefile --windowed --name "Seguimiento Ordenes" --hidden-import python_calamine app_seguimiento.py`

## Requisitos (una sola vez)
- Python 3.12 (instalado en `%LOCALAPPDATA%\Programs\Python\Python312`).
- Las librerías: `python -m pip install pandas xlsxwriter python-calamine pywin32`
- Microsoft Excel, que solo hace falta para la verificación.

Si `python` no se reconoce en la terminal, abre una terminal nueva o usa la ruta completa:
`%LOCALAPPDATA%\Programs\Python\Python312\python.exe`.

## Uso
```
python procesar_sap.py "Control de ordenes.csv" "Control de ordenes DD_MM.xlsx"
python verificar_libro.py "Control de ordenes DD_MM.xlsx" "Control de ordenes.csv"
```
- La entrada puede ser el **.csv** o el **.xlsx** que descarga SAP; los dos dan el mismo resultado.
- `procesar_sap.py` genera el libro y muestra en la consola el informe de limpieza. El mismo informe queda en la hoja **Log limpieza**.
- `verificar_libro.py` abre el libro en Excel y hace tres cosas:
  - comprueba que no haya errores de fórmula y que los controles den 0;
  - compara 6 combinaciones de filtros contra un cálculo independiente en pandas;
  - guarda el libro ya calculado.
  
  Termina con `RESULTADO: TODO OK` o con la lista de diferencias.

## Qué limpia
1. Quita el título, los encabezados repetidos y el pie "Impreso por SAP Business One".
2. Elimina las facturas anuladas: **Tipo = AN** o **Estado Factura = Anulado**. Si una OC solo tenía facturas anuladas, se conserva como fila *sin factura* con una marca en la columna Observación.
3. Elimina los duplicados por modificación de la OC: filas idénticas que solo difieren en *Aprobación*. Se conserva la **aprobación más reciente**.
4. Convierte los tipos:
   - fechas reales en formato dd/mm/yyyy, y las vacías quedan realmente vacías;
   - importes con formato #,##0.00, sin cambiar su valor;
   - RUC como texto;
   - Preliminar y Numero OC como enteros.
5. Valida los valores esperados (Estado, Serie, Mon, Tipo, Estado Factura) y la unicidad de RUC + N° Factura. Lo que encuentre se avisa en el Log, sin corregirlo.

## Columnas añadidas a Sheet1
| Col. | Contenido |
|---|---|
| V – Observación | Marcas para revisar y filtrar: fila sin factura en una OC que sí tiene facturas · OC cuya única factura fue anulada · estado de factura sin N° · factura sin Tipo · RUC con más de un nombre |
| W – Cuenta OC (filtros) | Vale 1 en la primera fila de cada OC que pasa los filtros. Sirve para contar cada orden una sola vez |
| X – Monto OC (filtros) | Total Sin IGV contado una sola vez por OC, en esa misma fila |

**Por qué W y X:** una OC con varias facturas repite su *Total Sin IGV* en cada fila. Las tablas de monto y de órdenes suman X y W, así que cada OC cuenta una sola vez aunque un filtro de factura (estado, emisión o pago) deje fuera alguna de sus filas. Las tablas de facturas cuentan filas con N° Factura.

## Uso del libro
- **Resumen Gastos:** los filtros están en J:O. Elige Sí/No en cada opción. Las fechas son opcionales; vacío significa sin filtro. Puedes elegir hasta 20 proveedores en el desplegable, y vacío significa todos.
- **Gráficos Proveedores:** elige la moneda en B4. Los gráficos y el Top 15 aplican los mismos filtros.
- Las celdas **"Control (debe ser 0)"** deben dar 0. Si no, algo cambió en los datos o en las fórmulas.
- Cuando aparezca un año nuevo o un usuario nuevo, basta con volver a ejecutar el script: las tablas se arman con los años y los usuarios que haya en los datos.
- Las fórmulas cubren las filas 3 a 10000 de Sheet1.

## Pendiente de tu decisión
Hay 6 RUC con más de un nombre de proveedor; el detalle y una propuesta están en Log limpieza. **No se han normalizado**: los filtros por proveedor y el Top 15 van por nombre, así que hoy cada variante aparece por separado.
