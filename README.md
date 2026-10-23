# Control de órdenes: procesamiento del reporte SAP

## Requerimientos e indicadores (pestañas «Requerimientos» e «Indicadores»)
Reemplaza el llenado manual del Excel «SEGUIMIENTO DE REQUERIMIENTOS Y ORDENES». La app es la fuente; el Excel se exporta.

- **➕ Nuevo requerimiento** abre un formulario con dos ayudas:
  - **Proyecto** con autocompletado por código o nombre (catálogo Consulta1). Al elegirlo muestra el controller, el controller 2 y el jefe de proyecto.
  - **Proveedor** con autocompletado (hoja PROVEEDORES + proveedores de SAP).
- Las fechas de cada flujo (valorización, orden, pago) tienen un botón **Hoy**.
- Al poner el **N° de orden**, con **⬇ Traer de SAP** se completan los campos vacíos: emisión y aprobación de OC, proveedor, factura y recepción (fecha de contabilización en SAP). Esto también ocurre solo en cada carga de un CSV de SAP. Reglas:
  - **Lo que escribiste no se toca.** Si vacías a propósito un dato que vino de SAP, no se vuelve a llenar.
  - **Lo que vino de SAP se mantiene al día**: si SAP cambia o deja de tener el dato, se actualiza, y queda en el historial.
  - **Monto y moneda** se toman siempre de lo aprobado en SAP y aparecen bloqueados con 🔗. **Excepción**: si la OC la comparten varios requerimientos, el total SAP es de la orden completa, así que cada uno registra su parte. El formulario avisa si la suma cuadra con SAP.
  - **Fecha de pago**: solo cuando todas las facturas están pagadas **y** la OC está cerrada o facturada al 100 %. Si no, es un pago parcial y la fecha no se llena.
  - Si cambias el N° de orden, los datos que habían venido de SAP se reemplazan por los de la nueva OC.
  - Junto a cada fecha se ve la fecha de SAP como referencia. Si la OC ya está en otro requerimiento, el formulario lo indica.
- La **etapa** se calcula sola según la fecha más avanzada registrada: Por valorizar → Valo por enviar a aprobación → Valo en aprobación → OC por emitir → OC por aprobar → OC por enviar → Pendiente de factura → Factura recibida → Factura en validación → Por enviar a pago → Pendiente de pago → Pagado.
  - ESTATUS = COMPLETO o ANULADO cierra el requerimiento.
  - Una OC cerrada o cancelada en SAP sin factura también lo cierra.
- La lista muestra por defecto los **pendientes hasta el envío de la OC**, resaltados. Filtros: etapa, tipo, controller, mes y búsqueda. Doble clic para editar. Cada cambio queda en el historial del requerimiento.
- **Indicadores**:
  - pendientes por etapa (N°, días promedio y máximo en la etapa, montos);
  - tiempos del flujo por tramo (N°, promedio, mediana, máximo).

  Se pueden filtrar por tipo, controller y rango de meses. Los días negativos (anticipos o regularizaciones: la fecha es anterior a la orden) no entran en los promedios y se cuentan aparte.
- **📤 Exportar Excel** genera `Seguimiento requerimientos DD_MM.xlsx` con tres hojas:
  - **CONTROL**: tus mismas columnas, más etapa, días, estado de la OC en SAP y último comentario de la OC;
  - **PENDIENTES**;
  - **TIEMPOS**: con fórmulas sobre la tabla.
- **📚 Actualizar catálogos** relee Consulta1 y PROVEEDORES del Excel de seguimiento (ruta guardada en la configuración). Los proyectos se reemplazan por la lista de Consulta1; los proveedores se suman a los que ya había.
- **📥 Importar desde Excel…** agrega, a nombre de tu comprador, los COD de la tabla CONTROL que todavía no importaste.
- Cada requerimiento tiene un **Comprador**. El filtro **Comprador** (arriba, junto a los botones, y en Indicadores) muestra por defecto los tuyos. **📤 Exportar Excel** exporta los del comprador elegido, con su nombre en el archivo.

## Aplicación «Seguimiento Ordenes.exe» (uso diario)
Para abrirla, haz doble clic en el `.exe`. No necesita Python. Tarda unos segundos en arrancar.

1. **📥 Cargar CSV de SAP…**: elige los dos reportes de SAP a la vez (con Ctrl). La app reconoce cuál es cuál:
   - **«Control de órdenes»** (CSV o xlsx): es el reporte general. Con él se arma el control del día y el seguimiento.
   - **Reporte de detalle «Seguimiento OC → Entrada → Factura → Pago»**: solo en **.xlsx**, porque el CSV de SAP sale en UTF-16 y desordena las columnas. Aporta las notas de entrada (ingreso de material), las facturas con su N° interno y «Pago parcial», y las líneas de cada OC con proyecto, unidad de negocio y área. **Solo se toman las OC aprobadas.**
   - Si cargas solo el detalle, la app vuelve a aplicar el último Control con la información logística nueva.
   - Si el reporte trae muchas menos órdenes que el anterior (reporte filtrado o incompleto), la app **pregunta antes de aplicarlo**.
   - Un reporte **anterior** a la última carga no modifica nada ni genera un control nuevo.

   La aplicación hace cuatro cosas:
   - limpia la data;
   - genera `Control de ordenes DD_MM.xlsx` en esta carpeta, con el mismo libro de siempre más la hoja **Seguimiento**;
   - guarda una copia del CSV en `Cargas\` para conservar el historial;
   - actualiza el seguimiento:
     - solo entran las **OC sin factura que están En proceso**. Las Cerradas o Canceladas sin factura ya no se van a facturar y no se tratan;
     - **excepción, OC «Cerrado» con material ingresado** (nota de entrada, según el reporte de detalle): ya se atendieron y solo falta ingresar la factura. Entran al seguimiento y se ven como «Cerrado (nota de entrada)». El requerimiento queda en «Pendiente de factura». En Power BI aparecen como «Pendiente de factura (nota de entrada)» y no cuentan como incumplimiento.
       - Si una OC todavía no está en el detalle, se usa una regla de respaldo: OC de bienes (OCB / 26OCB) con fecha desde el 01/09/2026 (`NOTA_ENTRADA_DESDE` en `seguimiento.py`).
       - Una OC cerrada **sin** ingreso de material sale del seguimiento como «OC cerrada sin factura».
     - los documentos de compra sin N° de factura (por ejemplo un DI, documento interno) cuentan como factura y se muestran como «DI s/n (OC …)»;
     - si en el nuevo reporte la OC ya trae factura, se cierra como **Facturada** y deja anotado el N° de factura;
     - si la OC pasa a Cerrada o Cancelada en SAP sin factura, se cierra como **OC cerrada sin factura** u **OC cancelada sin factura**, conservando sus comentarios;
     - si la OC ya no aparece en el reporte, se cierra como **Fuera del reporte**;
     - si la factura se anula, la OC vuelve a seguimiento;
     - los cambios de Estado de la OC quedan registrados.
2. **Seguimiento**: en la lista, elige una OC y completa:
   - estado de seguimiento, responsable y fecha compromiso;
   - **comentarios**, con **Agregar comentario** o Ctrl+Enter.

   Lo que editas se guarda solo al cambiar de OC, al comentar o al cambiar de pestaña; **Guardar cambios** sigue disponible. Un comentario a medio escribir queda como borrador de **esa** OC. Todo queda en el **Historial** con fecha y autor, y no se borra.
3. **Pestaña «Facturas»**: muestra todas las facturas del último reporte, una fila por factura.
   - Por defecto filtra **Pendiente de pago**. También puedes ver Pagado, Todos o (sin estado).
   - Otros filtros: comprador, estado de la OC, moneda, búsqueda, solo vencidas sin pagar y solo con comentarios.
   - Columnas: importes, emisión, vencimiento, **días para vencer** (negativo = vencida, en rojo) y fecha de pago. Las pagadas se ven en verde.
   - El resumen muestra el **importe pendiente por moneda** de lo que estás viendo.
   - Puedes comentar la OC de la factura seleccionada: el historial es único por OC.
   - En cada carga se registra en el historial de la OC cuando una factura **se paga** (con su fecha de pago), cuando aparece una **factura nueva** y cuando una factura **desaparece** del reporte (posible anulación).
4. **🔄 Actualizar Excel con comentarios**: vuelve a generar el control del último reporte con los comentarios del momento. Si el Excel está abierto, se guarda una copia con la hora en el nombre.
5. **Pestaña «🔎 Consultar OC»** (o **Ctrl+F** desde cualquier pestaña): el status de **cualquier** orden, sin importar en qué pestaña esté ni los filtros.
   - **Ver: Pendientes de ingreso** (o solo *Pendiente de atención* / *Atención parcial*) lista las OC de bienes abiertas con material por ingresar, de la más antigua a la más reciente. La búsqueda filtra esa lista. Al elegir una OC, el status muestra **cada línea** con su estado (ingresada, atendida por factura, ingreso parcial o pendiente de ingreso).
   - Busca por N° de OC o de factura, RUC, proveedor, o COD / referencia del requerimiento (por ejemplo «SOLPED 282»).
   - Busca entre todas las órdenes del último reporte SAP, más las de seguimiento, facturas y requerimientos, y es de cualquier comprador.
   - El panel **Status de la orden** reúne:
     - los datos de SAP;
     - un status en una línea: pagada, pendiente de pago (con las vencidas), pago parcial, en proceso sin factura, cerrada o cancelada sin factura;
     - cada factura con su vencimiento;
     - los requerimientos y su etapa;
     - el seguimiento y el último comentario.
   - **📋 Copiar status** lo deja listo para pegar en un correo o chat.
   - **Ver en OC sin factura**, **Ver en Facturas** y **Abrir requerimiento** te llevan a la orden en su pestaña, sin filtros que la oculten.
   - Abajo se ven su historial y el cuadro de comentario, aunque la orden no esté en «OC sin factura».

Filtros disponibles: texto, comprador, estado OC, situación, estado de seguimiento, **logística**, solo sin comentarios y solo compromisos vencidos.

La columna **Logística** viene del reporte de detalle y mide la **atención** de las OC de bienes. Como el registro de notas de entrada es reciente, **la factura también cuenta**:
- **Línea**:
  - *Ingresada*: tiene nota de entrada por toda la cantidad.
  - *Atendida por factura*: está facturada al 100 %, aunque no tenga nota.
  - *Ingreso parcial*.
  - *Pendiente de ingreso*: sin nota de entrada ni factura.
- **OC**:
  - **Atendida**: todas sus líneas están atendidas, o la OC está facturada por completo.
  - **Atención parcial**, con su %.
  - **Pendiente de atención**.
  - Si SAP ya la cerró: *Cerrada con atención parcial* o *Cerrada sin ingreso*.
  - Las OC de servicios figuran como «Servicio» y se siguen por factura.

Con eso se distingue la OC que espera el material de la que espera la factura. En Facturas, las pagadas en parte se ven como **Pago parcial**. En los requerimientos, la fecha de la última nota de entrada llena sola la **conformidad** de los bienes recibidos por completo, y se usa para el OTIF.

Colores de la lista:
- amarillo: OC sin comentarios;
- rojo: compromiso vencido;
- verde: OC facturada.

Si eliges el estado **"Cerrar sin factura"**, la OC deja de estar activa. Úsalo cuando ya sabes que una OC en proceso no se va a facturar, aunque todavía no esté cerrada en SAP.

**🚦 Semáforo de plazos**: cada pendiente se colorea según su plazo y tiene una columna **Plazo** y un filtro.

| Lista | Qué se mide | Amarillo (atención) | Rojo (atrasado) |
|---|---|---|---|
| Requerimientos | Días en la etapa actual | Según la etapa (ej. «Valo en aprobación»: 2 días) | Según la etapa (ej. 5 días) |
| OC sin factura | Días desde la fecha de la OC | 30 días | 60 días, o compromiso vencido |
| Facturas pendientes | Días al vencimiento | Vence en 7 días o menos | Vencida |

- Los umbrales se cambian en **⚙ Plazos…**. Los valores iniciales salen de tus tiempos actuales: de la solicitud al envío de la OC la mediana es 4 días.
- **🚦 Pendientes de hoy** se abre al iniciar la app (se puede desactivar). Muestra cuántos pendientes están atrasados y cuántos en atención en cada lista, y lo más atrasado hasta el envío de la OC. **Ver →** te lleva a la lista ya filtrada.
- En **Indicadores** y en el Excel exportado (columna PLAZO y hoja PENDIENTES) aparecen los atrasados y en atención por etapa.
- En «OC sin factura», las OC sin comentarios ahora se ven en **negrita**: el fondo de color queda para el semáforo.

**F5** actualiza la vista. Si el otro comprador o una carga programada cambió la base mientras la app estaba abierta, la vista se refresca sola al volver a la ventana.

### Power BI
La app deja al día **`Datos Power BI.xlsx`** en esta carpeta, con 5 tablas:
- **Ordenes**: todas las OC, con Categoría, Status, facturado, pendiente, seguimiento y plazo.
- **Facturas**: con días para vencer y plazo.
- **Requerimientos**: con etapa, plazo, fechas del flujo y días de cada tramo.
- **Comentarios**.
- **Info**: fecha del reporte SAP y de la exportación.

El archivo se regenera en cada carga del reporte SAP, al cerrar la app y con **📈 Datos para Power BI** en la pestaña Indicadores.

**Conectarlo (una sola vez):**
1. En Power BI Desktop, ve a **Obtener datos → Libro de Excel** y elige `Datos Power BI.xlsx`.
2. En el navegador, marca las **tablas** (icono de tabla, no las hojas): Ordenes, Facturas, Requerimientos, Comentarios e Info. Pulsa **Cargar**.
3. En la vista **Modelo**, revisa estas relaciones (Power BI suele crearlas solo); todas son de 1 a varios y de una sola dirección:
   - Ordenes[OC] → Facturas[OC];
   - Ordenes[OC] → Comentarios[OC];
   - Ordenes[OC] → Requerimientos[Orden].
4. Guarda el informe (`.pbix`) donde quieras.

Para ver los datos nuevos, pulsa **Inicio → Actualizar** en Power BI. El Excel debe estar cerrado al actualizar.

**Medidas útiles** (Nueva medida). Los montos están en distintas monedas: usa la columna **Moneda** como filtro o en la leyenda para no sumar soles con dólares.
```
OC = DISTINCTCOUNT(Ordenes[OC])
Pendiente de pago = SUM(Facturas[Importe pendiente])
Facturas vencidas = CALCULATE(COUNTROWS(Facturas), Facturas[Plazo factura] = "Vencida")
Requerimientos abiertos = CALCULATE(COUNTROWS(Requerimientos), Requerimientos[Abierto / cerrado] = "Abierto")
Atrasados = CALCULATE(COUNTROWS(Requerimientos), Requerimientos[Plazo] = "Atrasado")
Días solicitud → envío OC = AVERAGE(Requerimientos[Días: Solicitud → Envío de OC (total)])
```
El nombre exacto de cada columna «Días: …» se ve en la tabla Requerimientos. Los días negativos (anticipos o regularizaciones) vienen vacíos y se cuentan en «Tramos con anticipo».

**Visuales sugeridos:**
- Ordenes por **Categoría**, en anillo o barras, con segmentadores de Comprador, Moneda y Mes OC.
- **Pendiente de pago** por Proveedor (top 10) y por Mes vencimiento.
- **Requerimientos por Etapa**, con el Plazo como leyenda (atrasado, atención, en plazo).
- **Tiempos del flujo**: promedio de cada columna «Días: …» por Tipo o Controller.
- Una tabla de las OC «En proceso sin factura» con Días desde la OC, Responsable y Último comentario.

**Página «Logística»** (tabla Lineas y columnas nuevas de Ordenes):
- OC pendientes de ingreso (pendientes de atención y con atención parcial), **líneas pendientes de ingreso** y OC ingresadas sin factura;
- días promedio de la OC a la aprobación, de la OC a la nota de entrada y de la factura al pago;
- filtros por unidad de negocio, proyecto y comprador;
- la tabla de **líneas pendientes de ingreso** de las OC abiertas, con descripción, cantidad pedida e ingresada.

**KPIs de abastecimiento** (página «KPIs de abastecimiento» y tabla Proveedores):

| KPI | Cálculo | Dato que lo alimenta |
|---|---|---|
| 1. Ahorro | Monto referencial − monto de la OC (% sobre el referencial) | **Monto referencial** y **Tipo de referencia** en el formulario del requerimiento |
| 2. OTIF | A tiempo (conformidad ≤ fecha fin) **y** completo (OC facturada al 100 %) | **Conformidad (entrega real)** en el flujo de pago del requerimiento |
| 3. Lead time | Días de la solicitud al envío de la OC, a la factura, a la entrega (conformidad) y al pago | Fechas del requerimiento |
| 4. Cumplimiento | % de OC cerradas facturadas por el monto aprobado (±0.5 %) | Reporte SAP |
| 5. Desempeño de proveedores | Puntaje = 40 % cumplimiento + 30 % puntualidad (1.ª factura ≤ 30 días) + 30 % OTIF. Solo cuentan los criterios con 3 casos o más. A ≥ 85, B 70 a 85, C < 70 | Reporte SAP y requerimientos |

Los pesos, la tolerancia y los días se cambian al inicio de `powerbi.py` (`PESOS`, `TOLERANCIA`, `DIAS_FACTURA`, `MIN_EVAL`). Ahorro y OTIF muestran «Sin datos aún» hasta que registres los campos nuevos.

**Informe listo (`powerbi\Seguimiento.pbip`):**
- `python powerbi\actualizar_modelo.py` agrega al modelo las columnas, la tabla Proveedores y las medidas de los KPIs. Si lo vuelves a ejecutar, no duplica nada.
- `powerbi\tema_seguimiento.json` es el tema de colores (paleta, tarjetas con borde redondeado y sombra, encabezados de tabla). El generador lo aplica.
- Las medidas están en `powerbi\medidas.tmdl`. Se pegan en la vista TMDL de Power BI y se pulsa Aplicar.
- Las páginas **Resumen, Facturas, Requerimientos y Tiempos** se generan con `python powerbi\generar_paginas.py`, **con Power BI cerrado**.
  - El script revisa que cada campo exista en el modelo antes de escribir.
  - Si vuelves a ejecutarlo, regenera esas 4 páginas y no toca las que crees a mano.
- En el modelo, los montos son decimales con formato `#,0.00`, las fechas usan `dd/mm/yyyy` y las columnas «Mes …» usan `mmm yyyy`.
- El proyecto de Power BI no se sube a Git: tiene rutas locales y una caché con datos.

Para publicarlo en Power BI Service y compartirlo hace falta una licencia Pro. Como el archivo está en OneDrive, la actualización programada se puede configurar conectando por la ruta de SharePoint o OneDrive.

### Dos compradores en la misma base
- La carpeta de la aplicación debe estar en una **carpeta de red** (servidor), con el `.exe`, `seguimiento.db`, `Cargas\` y `Respaldos\`. Cada comprador abre el mismo `.exe` (conviene un acceso directo en el escritorio).
- **No uses OneDrive ni SharePoint para la base compartida.** Cada equipo sincroniza su propia copia y, si los dos guardan, OneDrive crea copias en conflicto y se pierden cambios.
- La primera vez, la app pregunta **¿Quién usa la aplicación?**: el comprador (tal como aparece en la columna «Usuario» de SAP) y la firma de los comentarios. Se cambia con **Cambiar**, arriba a la derecha.
- Requerimientos, OC sin factura, Facturas, Indicadores y **Pendientes de hoy** muestran por defecto **lo tuyo**. Con **Comprador: Todos** ves también lo del otro.
- Los comentarios de una OC son compartidos: los dos ven el mismo historial, con la firma de cada uno.
- Si el otro guarda mientras tú haces lo mismo, la app espera unos segundos y luego guarda. Si los dos editan el **mismo requerimiento**, la app avisa al guardar y conserva solo los campos que tú cambiaste; el resto queda como lo dejó el otro.
- El reporte SAP lo puede cargar cualquiera de los dos; trae las OC de ambos. Si el otro ya cargó el mismo reporte, la segunda carga no cambia nada.
- **Importar el Excel de otro comprador**: entra con su usuario y usa **📥 Importar desde Excel…**. Sus COD se numeran a continuación de los existentes. El COD original se ve en el título del formulario y sirve para buscar.

Archivos de la aplicación (deben quedarse junto al `.exe`):
- `seguimiento.db`: requerimientos, comentarios e historial. **No lo borres.**
- `Cargas\`: copias de los reportes cargados.
- `Respaldos\`: copias de seguridad de la base. En `Respaldos\Automáticos` hay una copia diaria automática, de los últimos 15 días. Se hace al abrir la app o al cargar un reporte.
- `errores.log`: detalle técnico de cualquier error (para revisarlo si algo falla).

La configuración de cada usuario (comprador, firma, rutas y plazos) está en `%APPDATA%\Seguimiento Ordenes\config.json`, no en la carpeta compartida.

**Carga automática (opcional)**: con el Programador de tareas de Windows puedes ejecutar
`"Seguimiento Ordenes.exe" --cargar "C:\...\Control de ordenes.csv" "C:\...\Seguimiento.xlsx"`. Puedes pasar uno o los dos reportes. El resultado queda en `ultima_carga.log`, que empieza con uno de estos estados:
- **OK**: la carga se aplicó.
- **OMITIDO**: el reporte era anterior a la última carga.
- **REVISAR**: el reporte parecía incompleto y no se aplicó. Cárgalo desde la app para confirmarlo, o agrega `--forzar`.
- **ERROR**: algo falló; el detalle está en el mismo log.

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
