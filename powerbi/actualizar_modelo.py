"""
Pone al día el modelo de Power BI (Seguimiento.SemanticModel) con las columnas y medidas de los KPI.

    python actualizar_modelo.py

- Agrega a Ordenes y Requerimientos las columnas nuevas de «Datos Power BI.xlsx» (con su tipo en Power Query).
- Crea las tablas Proveedores (scorecard) y Lineas (detalle de cada OC), con sus relaciones.
- Agrega las medidas de la carpeta «KPIs de abastecimiento» a la tabla Medidas.
Es idempotente: lo que ya existe no se toca. Power BI Desktop debe estar CERRADO.
Al abrir el informe, pulsa «Actualizar» para cargar las columnas nuevas.
"""
import os
import re
import sys
import uuid

AQUI = os.path.dirname(os.path.abspath(__file__))
DEF = os.path.join(AQUI, "Seguimiento.SemanticModel", "definition")
TABLAS = os.path.join(DEF, "tables")

# (nombre, tipo) — tipo: texto, entero, decimal, porcentaje, fecha
COLUMNAS = {
    "Ordenes": [("Cumplimiento", "texto"), ("Cumplimiento evaluado", "texto"), ("% desviación facturado", "porcentaje"),
                ("Días OC → 1ra factura", "entero"), ("Factura a tiempo", "texto"),
                ("Tipo OC", "texto"), ("Estado logístico", "texto"), ("Fecha última entrada", "fecha"),
                ("% ingresado", "porcentaje"), ("Proyecto", "texto"), ("Unidad de negocio", "texto"), ("Área", "texto"),
                ("Días OC → aprobación", "entero"), ("Días OC → entrada", "entero"), ("Días factura → pago", "entero")],
    "Requerimientos": [("Monto referencial", "decimal"), ("Tipo de referencia", "texto"), ("Ahorro", "decimal"),
                       ("Base del ahorro", "decimal"), ("% ahorro", "porcentaje"), ("Fecha de conformidad", "fecha"),
                       ("A tiempo", "texto"), ("Completo", "texto"), ("OTIF", "texto"), ("Lead time (días)", "entero")],
    "Proveedores": [("RUC", "texto"), ("Proveedor", "texto"), ("Último comprador", "texto"), ("N° OC", "entero"),
                    ("Última OC", "fecha"), ("OC evaluadas", "entero"), ("% cumplimiento", "porcentaje"),
                    ("OC con factura", "entero"), ("% puntualidad", "porcentaje"),
                    ("Días promedio OC → factura", "decimal"), ("Entregas evaluadas (OTIF)", "entero"),
                    ("% OTIF", "porcentaje"), ("OC sin factura atrasadas", "entero"), ("Puntaje", "decimal"),
                    ("Clasificación", "texto"), ("Criterios usados", "texto")],
    "Lineas": [("OC", "entero"), ("Línea", "entero"), ("Tipo OC", "texto"), ("Código", "texto"),
               ("Descripción", "texto"), ("Cantidad", "decimal"), ("UM", "texto"), ("Total línea sin IGV", "decimal"),
               ("Moneda", "texto"), ("Proyecto", "texto"), ("Unidad de negocio", "texto"), ("Área", "texto"),
               ("Cantidad ingresada", "decimal"), ("Pendiente de ingreso", "decimal"), ("Estado logístico", "texto"),
               ("Fecha de entrada", "fecha")],
}
# Tablas nuevas: (columna en la tabla, tabla y columna del lado «uno» de la relación)
TABLAS_NUEVAS = {"Proveedores": ("Ordenes.RUC", "Proveedores.RUC"), "Lineas": ("Lineas.OC", "Ordenes.OC")}
TIPOS = {  # tipo -> (dataType, formatString, summarizeBy, tipo M)
    "texto": ("string", None, "none", "type text"),
    "entero": ("int64", "0", "sum", "Int64.Type"),
    "decimal": ("double", "#,0.00", "sum", "type number"),
    "porcentaje": ("double", "0.0%", "none", "type number"),
    "fecha": ("dateTime", "dd/mm/yyyy", "none", "type date"),
}
CARPETA = "KPIs de abastecimiento"
MEDIDAS = [
    # 1. Ahorro en compras
    ("Ahorro total", "COALESCE(SUM(Requerimientos[Ahorro]), 0)", "#,0.00"),
    ("% Ahorro sobre referencia", "DIVIDE(SUM(Requerimientos[Ahorro]), SUM(Requerimientos[Base del ahorro]))", "0.0%"),
    ("Requerimientos con referencia", "COALESCE(COUNT(Requerimientos[Ahorro]), 0)", "#,0"),
    # 2. OTIF
    ("OTIF %", 'DIVIDE(CALCULATE(COUNTROWS(Requerimientos), Requerimientos[OTIF] = "Sí"), '
               'CALCULATE(COUNTROWS(Requerimientos), Requerimientos[OTIF] IN {"Sí", "No"}))', "0.0%"),
    ("% A tiempo", 'DIVIDE(CALCULATE(COUNTROWS(Requerimientos), Requerimientos[A tiempo] = "Sí"), '
                   'CALCULATE(COUNTROWS(Requerimientos), Requerimientos[A tiempo] IN {"Sí", "No"}))', "0.0%"),
    ("% Completo", 'DIVIDE(CALCULATE(COUNTROWS(Requerimientos), Requerimientos[Completo] = "Sí"), '
                   'CALCULATE(COUNTROWS(Requerimientos), Requerimientos[Completo] IN {"Sí", "No"}))', "0.0%"),
    ("Entregas evaluadas", 'COALESCE(CALCULATE(COUNTROWS(Requerimientos), Requerimientos[OTIF] IN {"Sí", "No"}), 0)', "#,0"),
    # 3. Lead time
    ("Lead time a la entrega", "AVERAGE(Requerimientos[Lead time (días)])", "0.0"),
    # 4. Cumplimiento de la compra
    ("Cumplimiento de la compra", 'DIVIDE(CALCULATE([N° OC], Ordenes[Cumplimiento] = "Cumple"), '
                                  'CALCULATE([N° OC], Ordenes[Cumplimiento evaluado] = "Sí"))', "0.0%"),
    ("OC evaluadas (cumplimiento)", 'CALCULATE([N° OC], Ordenes[Cumplimiento evaluado] = "Sí")', "#,0"),
    ("Puntualidad de facturación", 'DIVIDE(CALCULATE([N° OC], Ordenes[Factura a tiempo] = "Sí"), '
                                   'CALCULATE([N° OC], Ordenes[Factura a tiempo] IN {"Sí", "No"}))', "0.0%"),
    # 5. Desempeño de proveedores
    ("Puntaje promedio proveedores", "AVERAGE(Proveedores[Puntaje])", "0.0"),
    ("Proveedores evaluados", "COALESCE(COUNT(Proveedores[Puntaje]), 0)", "#,0"),
    ("Proveedores categoría A", 'COALESCE(CALCULATE(COUNTROWS(Proveedores), LEFT(Proveedores[Clasificación], 1) = "A"), 0)', "#,0"),
    ("N° proveedores", "COUNTROWS(Proveedores)", "#,0"),
    # Para tarjetas: texto «Sin datos aún» mientras no haya registros (en vez de «(En blanco)»)
    ("Ahorro sobre referencia", 'VAR v = [% Ahorro sobre referencia] RETURN IF(ISBLANK(v), "Sin datos aún", FORMAT(v, "0.0%"))', None),
    ("OTIF (a tiempo y completo)", 'VAR v = [OTIF %] RETURN IF(ISBLANK(v), "Sin datos aún", FORMAT(v, "0.0%"))', None),
    ("Lead time a la entrega (días)", 'VAR v = [Lead time a la entrega] RETURN IF(ISBLANK(v), "Sin datos aún", FORMAT(v, "0.0"))', None),
    # Logística (reporte de detalle)
    ("OC pendientes de ingreso", 'CALCULATE([N° OC], Ordenes[Estado logístico] = "Pendiente de ingreso")', "#,0"),
    ("OC con ingreso parcial", 'CALCULATE([N° OC], Ordenes[Estado logístico] = "Ingreso parcial")', "#,0"),
    ("OC ingresadas sin factura", 'CALCULATE([N° OC], Ordenes[Categoría] = "Pendiente de factura (nota de entrada)")', "#,0"),
    ("Prom. días OC → aprobación", "AVERAGE(Ordenes[Días OC → aprobación])", "0.0"),
    ("Prom. días OC → entrada", "AVERAGE(Ordenes[Días OC → entrada])", "0.0"),
    ("Prom. días factura → pago", "AVERAGE(Ordenes[Días factura → pago])", "0.0"),
    ("Monto de líneas", "SUM(Lineas[Total línea sin IGV])", "#,0.00"),
]


def _q(nombre):
    return f"'{nombre}'" if re.search(r"[^A-Za-z0-9_]", nombre) else nombre


def _columna(nombre, tipo):
    dt, fmt, resumen, _ = TIPOS[tipo]
    lineas = [f"\tcolumn {_q(nombre)}", f"\t\tdataType: {dt}"]
    if fmt:
        lineas.append(f"\t\tformatString: {fmt}")
    lineas += [f"\t\tlineageTag: {uuid.uuid4()}", f"\t\tsummarizeBy: {resumen}", f"\t\tsourceColumn: {nombre}", "",
               "\t\tannotation SummarizationSetBy = Automatic", ""]
    if tipo == "fecha":
        lineas.insert(-1, "\t\tannotation UnderlyingDateTimeDataType = Date")
        lineas.insert(-1, "")
    return "\n".join(lineas) + "\n"


def _columnas_existentes(txt):
    return {m.group(2) for m in re.finditer(r"^\tcolumn\s+('?)(.+?)\1\s*$", txt, re.M)}


def _ruta_excel():
    txt = open(os.path.join(TABLAS, "Ordenes.tmdl"), encoding="utf-8").read()
    return re.search(r'File\.Contents\("([^"]+)"\)', txt).group(1)


def columnas_en_tabla(tabla):
    ruta = os.path.join(TABLAS, f"{tabla}.tmdl")
    txt = open(ruta, encoding="utf-8").read()
    nuevas = [(n, t) for n, t in COLUMNAS[tabla] if n not in _columnas_existentes(txt)]
    if not nuevas:
        return 0
    i = txt.index("\tpartition ")
    txt = txt[:i] + "".join(_columna(n, t) for n, t in nuevas) + txt[i:]
    # Tipos en el paso «Tipo cambiado» de Power Query
    m = re.search(r'(#"Tipo cambiado" = Table\.TransformColumnTypes\(\w+,\{.*)\}\)', txt)
    extra = "".join(f', {{"{n}", {TIPOS[t][3]}}}' for n, t in nuevas)
    txt = txt[:m.end(1)] + extra + txt[m.end(1):]
    open(ruta, "w", encoding="utf-8", newline="").write(txt)
    return len(nuevas)


def tabla_nueva(nombre):
    """Crea la tabla `nombre` (Excel «Datos Power BI.xlsx») con su relación; si ya existe, agrega columnas faltantes."""
    ruta = os.path.join(TABLAS, f"{nombre}.tmdl")
    if os.path.exists(ruta):
        return columnas_en_tabla(nombre)
    tipos = ", ".join(f'{{"{n}", {TIPOS[t][3]}}}' for n, t in COLUMNAS[nombre])
    txt = (f"table {nombre}\n\tlineageTag: {uuid.uuid4()}\n\n"
           + "".join(_columna(n, t) for n, t in COLUMNAS[nombre])
           + f"\tpartition {nombre} = m\n\t\tmode: import\n\t\tsource =\n\t\t\t\tlet\n"
           + f'\t\t\t\t    Origen = Excel.Workbook(File.Contents("{_ruta_excel()}"), null, true),\n'
           + f'\t\t\t\t    {nombre}_Table = Origen{{[Item="{nombre}",Kind="Table"]}}[Data],\n'
           + f'\t\t\t\t    #"Tipo cambiado" = Table.TransformColumnTypes({nombre}_Table,{{{tipos}}})\n'
           + '\t\t\t\tin\n\t\t\t\t    #"Tipo cambiado"\n\n'
           + "\tannotation PBI_ResultType = Table\n\n\tannotation PBI_NavigationStepName = Navegación\n\n")
    open(ruta, "w", encoding="utf-8", newline="").write(txt)
    modelo = os.path.join(DEF, "model.tmdl")
    m = open(modelo, encoding="utf-8").read()
    if f"ref table {nombre}\n" not in m + "\n":
        m = m.replace('annotation PBI_QueryOrder = [', f'annotation PBI_QueryOrder = ["{nombre}",', 1)
        m = m.rstrip("\n") + f"\nref table {nombre}\n"
        open(modelo, "w", encoding="utf-8", newline="").write(m)
    desde, hacia = TABLAS_NUEVAS[nombre]
    rel = os.path.join(DEF, "relationships.tmdl")
    r = open(rel, encoding="utf-8").read()
    if f"fromColumn: {desde}\n\ttoColumn: {hacia}" not in r:
        r = r.rstrip("\n") + f"\n\nrelationship {uuid.uuid4()}\n\tfromColumn: {desde}\n\ttoColumn: {hacia}\n"
        open(rel, "w", encoding="utf-8", newline="").write(r)
    return len(COLUMNAS[nombre])


def medidas():
    ruta = os.path.join(TABLAS, "Medidas.tmdl")
    txt = open(ruta, encoding="utf-8").read()
    existentes = {m.group(2) for m in re.finditer(r"^\tmeasure\s+('?)(.+?)\1\s*=", txt, re.M)}
    columnas = set()
    for f in os.listdir(TABLAS):
        columnas |= _columnas_existentes(open(os.path.join(TABLAS, f), encoding="utf-8").read())
    choque = [n for n, _, _ in MEDIDAS if n in columnas]
    if choque:
        sys.exit(f"Medidas con el mismo nombre que una columna: {choque}")
    nuevas = [(n, e, f) for n, e, f in MEDIDAS if n not in existentes]
    bloque = "".join(f"\tmeasure {_q(n)} = {e}\n" + (f"\t\tformatString: {f}\n" if f else "")
                     + f"\t\tdisplayFolder: {CARPETA}\n\t\tlineageTag: {uuid.uuid4()}\n\n" for n, e, f in nuevas)
    i = txt.index("\tcolumn ")
    open(ruta, "w", encoding="utf-8", newline="").write(txt[:i] + bloque + txt[i:])
    return len(nuevas)


def main():
    print("Ordenes: columnas nuevas", columnas_en_tabla("Ordenes"))
    print("Requerimientos: columnas nuevas", columnas_en_tabla("Requerimientos"))
    for nombre in TABLAS_NUEVAS:
        print(f"{nombre}: columnas", tabla_nueva(nombre))
    print("Medidas nuevas", medidas())
    print("Listo. Abre Seguimiento.pbip y pulsa «Actualizar».")


if __name__ == "__main__":
    main()
