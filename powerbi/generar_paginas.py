"""
Genera las páginas del informe de Power BI (proyecto Seguimiento.pbip, formato PBIR).

    python generar_paginas.py

- Crea o reemplaza las páginas «KPIs de abastecimiento», Resumen, Facturas, Requerimientos y Tiempos.
  Las páginas creadas a mano en Power BI no se tocan.
- Aplica el tema de colores tema_seguimiento.json a todo el informe.
- Antes de escribir, comprueba que cada tabla, columna y medida exista en el modelo
  (Seguimiento.SemanticModel). Si falta alguna, no escribe nada (ejecuta antes actualizar_modelo.py).
- Power BI Desktop debe estar CERRADO: si está abierto, al guardar sobrescribe estos archivos.
"""
import hashlib
import json
import os
import re
import shutil
import sys

AQUI = os.path.dirname(os.path.abspath(__file__))
REPORTE = os.path.join(AQUI, "Seguimiento.Report", "definition")
RECURSOS = os.path.join(AQUI, "Seguimiento.Report", "StaticResources", "RegisteredResources")
MODELO = os.path.join(AQUI, "Seguimiento.SemanticModel", "definition", "tables")
TEMA = os.path.join(AQUI, "tema_seguimiento.json")
NOMBRE_TEMA = "TemaSeguimiento.json"
ESQ = "https://developer.microsoft.com/json-schemas/fabric/item/report/definition"
ANCHO, ALTO = 1280, 720
M = 8                                           # margen

# Distribución estándar de una página (y, alto)
ENCABEZADO = (0, 50)
TARJETAS = (58, 92)
FILTROS = (158, 58)
GRAFICOS = (224, 244)
TABLA = (476, 236)

AZUL_OSCURO = "#1E3A8A"
ROJO, AMBAR, VERDE, GRIS = "#E11D48", "#F59E0B", "#16A34A", "#94A3B8"
ACENTOS = ["#2563EB", "#0EA5A4", "#E11D48", "#F59E0B", "#7C3AED", "#16A34A", "#EA580C"]
COLORES_PLAZO = {"Atrasado": ROJO, "Atención": AMBAR, "En plazo": VERDE}
COLORES_FACTURA = {"Vencida": ROJO, "Por vencer": AMBAR, "Al día": VERDE}
COLORES_CATEGORIA = {"Pagada": VERDE, "Pendiente de pago": AMBAR, "En proceso sin factura": "#EA580C",
                     "Pago parcial": "#60A5FA", "Cerrada sin factura": GRIS, "Cancelada sin factura": "#64748B",
                     "No está en el reporte": "#CBD5E1"}
COLORES_CUMPLIMIENTO = {"Cumple": VERDE, "Cerrada con saldo sin facturar": AMBAR, "Cerrada sin factura": GRIS,
                        "Sobrefacturada": ROJO}
COLORES_CLASIFICACION = {"A (≥ 85)": VERDE, "B (70 a 85)": AMBAR, "C (< 70)": ROJO}


def _id(*partes):
    return hashlib.md5("|".join(partes).encode("utf-8")).hexdigest()[:20]


def _lit(v):
    if isinstance(v, bool):
        return {"expr": {"Literal": {"Value": "true" if v else "false"}}}
    if isinstance(v, (int, float)):
        return {"expr": {"Literal": {"Value": f"{v}D"}}}
    return {"expr": {"Literal": {"Value": "'" + str(v).replace("'", "''") + "'"}}}


def _color(hex_):
    return {"solid": {"color": _lit(hex_)}}


# ---------------------------------------------------------------- modelo (para validar)
def leer_modelo():
    """{tabla: {'columnas': set, 'medidas': set}} leído de los .tmdl del modelo."""
    modelo = {}
    for f in os.listdir(MODELO):
        if not f.endswith(".tmdl") or f.startswith(("LocalDateTable", "DateTableTemplate")):
            continue
        txt = open(os.path.join(MODELO, f), encoding="utf-8").read()
        tabla = re.search(r"^table\s+('?)(.+?)\1\s*$", txt, re.M).group(2)
        nombre = lambda tipo: {m.group(2) for m in re.finditer(rf"^\t{tipo}\s+('?)(.+?)\1(\s*=.*)?$", txt, re.M)}
        modelo[tabla] = {"columnas": nombre("column"), "medidas": nombre("measure")}
    return modelo


USADOS = []          # (tabla, campo, es_medida) de todo lo que referencian las páginas


# ---------------------------------------------------------------- campos
def col(tabla, campo):
    USADOS.append((tabla, campo, False))
    return {"field": {"Column": {"Expression": {"SourceRef": {"Entity": tabla}}, "Property": campo}},
            "queryRef": f"{tabla}.{campo}", "nativeQueryRef": campo}


def med(nombre):
    USADOS.append(("Medidas", nombre, True))
    return {"field": {"Measure": {"Expression": {"SourceRef": {"Entity": "Medidas"}}, "Property": nombre}},
            "queryRef": f"Medidas.{nombre}", "nativeQueryRef": nombre}


def filtro_in(nombre, tabla, campo, valores):
    """Filtro de objeto visual: tabla[campo] en (valores)."""
    USADOS.append((tabla, campo, False))
    return {"name": _id("filtro", nombre), "type": "Categorical",
            "field": {"Column": {"Expression": {"SourceRef": {"Entity": tabla}}, "Property": campo}},
            "filter": {"Version": 2, "From": [{"Name": "t", "Entity": tabla, "Type": 0}],
                       "Where": [{"Condition": {"In": {
                           "Expressions": [{"Column": {"Expression": {"SourceRef": {"Source": "t"}}, "Property": campo}}],
                           "Values": [[{"Literal": {"Value": "'" + v + "'"}}] for v in valores]}}}]}}


def colores(tabla, campo, por_valor):
    """objects.dataPoint: un color fijo para cada valor de tabla[campo] (series o categorías)."""
    return {"dataPoint": [
        {"properties": {"fill": _color(color)},
         "selector": {"data": [{"scopeId": {"Comparison": {
             "ComparisonKind": 0,
             "Left": {"Column": {"Expression": {"SourceRef": {"Entity": tabla}}, "Property": campo}},
             "Right": {"Literal": {"Value": "'" + valor + "'"}}}}}]}}
        for valor, color in por_valor.items()]}


def colores_medidas(por_medida):
    """objects.dataPoint: un color por medida (gráficos con varias medidas)."""
    return {"dataPoint": [{"properties": {"fill": _color(color)}, "selector": {"metadata": f"Medidas.{m}"}}
                          for m, color in por_medida.items()]}


# ---------------------------------------------------------------- visuales
def visual(pagina, i, tipo, x, y, w, h, roles=None, titulo=None, subtitulo=None, objetos=None, filtros=None,
           orden=None, contenedor=None):
    nombre = _id(pagina, str(i))
    v = {"visualType": tipo, "drillFilterOtherVisuals": True}
    if roles:
        proyecciones = {}
        for rol, campos in roles.items():
            ps = []
            for k, c in enumerate(campos):
                p = dict(c)
                if k == 0 and rol in ("Category", "Values") and tipo != "tableEx":
                    p["active"] = True
                ps.append(p)
            proyecciones[rol] = {"projections": ps}
        v["query"] = {"queryState": proyecciones}
    if orden:
        campo, desc = orden
        v["query"]["sortDefinition"] = {"sort": [{"field": campo["field"], "direction": "Descending" if desc else "Ascending"}],
                                        "isDefaultSort": True}
    if objetos:
        v["objects"] = objetos
    vco = dict(contenedor or {})
    if titulo:
        vco["title"] = [{"properties": {"show": _lit(True), "text": _lit(titulo)}}]
    if subtitulo:
        vco["subTitle"] = [{"properties": {"show": _lit(True), "text": _lit(subtitulo)}}]
    if vco:
        v["visualContainerObjects"] = vco
    cont = {"$schema": f"{ESQ}/visualContainer/2.14.0/schema.json", "name": nombre,
            "position": {"x": x, "y": y, "z": i, "height": h, "width": w, "tabOrder": i}, "visual": v}
    if filtros:
        cont["filterConfig"] = {"filters": filtros}
    return cont


def encabezado(pagina, titulo, subtitulo, color=AZUL_OSCURO):
    """Banda superior con el título de la página (cuadro de texto con fondo de color)."""
    y, h = ENCABEZADO
    parrafos = [{"textRuns": [{"value": titulo, "textStyle": {"fontWeight": "bold", "fontSize": "18pt",
                                                                "color": "#FFFFFF"}},
                              {"value": "   " + subtitulo, "textStyle": {"fontSize": "10pt", "color": "#BFDBFE"}}]}]
    return visual(pagina, 99, "textbox", 0, y, ANCHO, h, objetos={"general": [{"properties": {"paragraphs": parrafos}}]},
                  contenedor={"background": [{"properties": {"show": _lit(True), "color": _color(color),
                                                             "transparency": _lit(0)}}],
                              "border": [{"properties": {"show": _lit(False)}}],
                              "dropShadow": [{"properties": {"show": _lit(False)}}]})


def tarjeta(pagina, i, medida, x, y, w, h, acento, tamano=None, etiqueta=10):
    """Tarjeta con el nombre arriba (título), el valor en el color de acento, fondo suave y borde de acento."""
    labels = {"color": _color(acento)}
    if tamano:
        labels["fontSize"] = _lit(tamano)
    return visual(pagina, i, "card", x, y, w, h, {"Values": [med(medida)]},
                  objetos={"labels": [{"properties": labels}],
                           "categoryLabels": [{"properties": {"show": _lit(False)}}]},
                  contenedor={"title": [{"properties": {"show": _lit(True), "text": _lit(medida),
                                                        "fontSize": _lit(etiqueta), "alignment": _lit("center"),
                                                        "fontColor": _color("#334155"), "bold": _lit(False)}}],
                              "subTitle": [{"properties": {"show": _lit(False)}}],
                              "background": [{"properties": {"show": _lit(True), "color": _color(acento),
                                                             "transparency": _lit(90)}}],
                              "border": [{"properties": {"show": _lit(True), "color": _color(acento),
                                                         "radius": _lit(12)}}]})


def tarjetas(pagina, inicio, medidas, y=TARJETAS[0], h=TARJETAS[1], acentos=ACENTOS, tamano=20, etiqueta=9):
    w = (ANCHO - 2 * M - M * (len(medidas) - 1)) / len(medidas)
    return [tarjeta(pagina, inicio + k, m, M + k * (w + M), y, w, h, acentos[k % len(acentos)], tamano, etiqueta)
            for k, m in enumerate(medidas)]


def segmentadores(pagina, inicio, campos, y=FILTROS[0], extra_medida="Datos al"):
    vs = []
    w = 230
    for k, (tabla, campo) in enumerate(campos):
        vs.append(visual(pagina, inicio + k, "slicer", M + k * (w + M), y, w, FILTROS[1], {"Values": [col(tabla, campo)]},
                         objetos={"data": [{"properties": {"mode": _lit("Dropdown")}}]}))
    if extra_medida:
        vs.append(visual(pagina, inicio + len(campos), "card", ANCHO - M - 300, y, 300, FILTROS[1],
                         {"Values": [med(extra_medida)]},
                         objetos={"labels": [{"properties": {"fontSize": _lit(12), "color": _color(AZUL_OSCURO)}}],
                                  "categoryLabels": [{"properties": {"show": _lit(False)}}]}))
    return vs


def fila(n, y=GRAFICOS[0], h=GRAFICOS[1]):
    """Posiciones (x, y, w, h) de n visuales lado a lado."""
    w = (ANCHO - 2 * M - M * (n - 1)) / n
    return [(M + k * (w + M), y, w, h) for k in range(n)]


# ---------------------------------------------------------------- páginas
def pagina_kpis():
    p = "kpis"
    vs = [encabezado(p, "KPIs de abastecimiento", "5 indicadores para una cadena de suministro más eficiente")]
    # 1 Ahorro · 2 OTIF · 3 Lead time · 4 Cumplimiento · 5 Desempeño (colores de cada indicador)
    colores_kpi = ["#2563EB", "#16A34A", "#7C3AED", "#EA580C", "#0891B2"]
    vs += tarjetas(p, 0, ["Ahorro sobre referencia", "OTIF (a tiempo y completo)", "Días solicitud → envío de OC",
                          "Cumplimiento de la compra", "Puntaje promedio proveedores"],
                   y=58, h=90, acentos=colores_kpi, tamano=22, etiqueta=10)
    vs += tarjetas(p, 10, ["Ahorro total", "Entregas evaluadas", "Lead time a la entrega (días)",
                           "Puntualidad de facturación", "Proveedores categoría A"],
                   y=154, h=80, acentos=colores_kpi, tamano=13, etiqueta=9)
    vs += segmentadores(p, 20, [("Ordenes", "Comprador"), ("Ordenes", "Mes OC")], y=240)
    y1, h1 = 304, 192
    a, b, c = fila(3, y1, h1)
    vs.append(visual(p, 30, "pieChart", *a, {"Category": [col("Ordenes", "Cumplimiento")], "Y": [med("N° OC")]},
                     titulo="4 · Cumplimiento de la compra",
                     subtitulo="OC cerradas facturadas por el monto aprobado (±0.5 %)",
                     objetos={**colores("Ordenes", "Cumplimiento", COLORES_CUMPLIMIENTO),
                              "labels": [{"properties": {"labelStyle": _lit("Percent of total"),
                                                         "labelDisplayUnits": _lit(1)}}]},
                     filtros=[filtro_in("fEval", "Ordenes", "Cumplimiento evaluado", ["Sí"])]))
    tramos = {"Días solicitud → envío de OC": "#7C3AED", "Días envío de OC → factura": "#A78BFA",
              "Lead time a la entrega": "#2563EB", "Días solicitud → pago": "#C4B5FD"}
    vs.append(visual(p, 31, "clusteredColumnChart", *b, {"Category": [col("Requerimientos", "Tipo")],
                                                          "Y": [med(m) for m in tramos]},
                     titulo="3 · Lead time del abastecimiento (días promedio)",
                     subtitulo="Por tipo, desde la solicitud hasta cada hito (entrega = fecha de conformidad)",
                     objetos={**colores_medidas(tramos), "labels": [{"properties": {"show": _lit(True)}}]}))
    vs.append(visual(p, 32, "lineChart", *c, {"Category": [col("Ordenes", "Mes OC")],
                                               "Y": [med("Cumplimiento de la compra"), med("Puntualidad de facturación")]},
                     titulo="Cumplimiento y puntualidad por mes",
                     subtitulo="% de OC que cumplen y que se facturan dentro de 30 días",
                     objetos=colores_medidas({"Cumplimiento de la compra": "#EA580C",
                                              "Puntualidad de facturación": "#0891B2"})))
    y2, h2 = 502, 210
    d, e = fila(2, y2, h2)
    d = (d[0], d[1], d[2] * 1.35, d[3])
    e = (d[0] + d[2] + M, e[1], ANCHO - M - (d[0] + d[2] + M), e[3])
    clasif = filtro_in("fClas", "Proveedores", "Clasificación", list(COLORES_CLASIFICACION))
    puntaje = col("Proveedores", "Puntaje")
    vs.append(visual(p, 33, "tableEx", *d, {"Values": [
        col("Proveedores", c) for c in ("Proveedor", "N° OC", "% cumplimiento", "% puntualidad", "% OTIF")] + [puntaje,
        col("Proveedores", "Clasificación")]},
        titulo="5 · Desempeño de proveedores (scorecard)",
        objetos={"columnWidth": [{"properties": {"value": _lit(ancho)}, "selector": {"metadata": f"Proveedores.{c}"}}
                                 for c, ancho in (("Proveedor", 230), ("N° OC", 55), ("% cumplimiento", 115),
                                                  ("% puntualidad", 105), ("% OTIF", 65), ("Puntaje", 70),
                                                  ("Clasificación", 100))]},
        subtitulo="Puntaje = 40 % cumplimiento + 30 % puntualidad + 30 % OTIF (solo criterios con 3 casos o más)",
        filtros=[clasif], orden=(puntaje, True)))
    vs.append(visual(p, 34, "donutChart", *e, {"Category": [col("Proveedores", "Clasificación")],
                                                "Y": [med("N° proveedores")]},
                     titulo="Proveedores por clasificación", subtitulo="A ≥ 85 · B 70 a 85 · C < 70",
                     objetos={**colores("Proveedores", "Clasificación", COLORES_CLASIFICACION),
                              "labels": [{"properties": {"labelStyle": _lit("Data value, percent of total"),
                                                         "labelDisplayUnits": _lit(1)}}]},
                     filtros=[filtro_in("fClas2", "Proveedores", "Clasificación", list(COLORES_CLASIFICACION))]))
    return "KPIs de abastecimiento", vs


def pagina_resumen():
    p = "resumen"
    vs = [encabezado(p, "Resumen", "Órdenes, facturas y requerimientos de un vistazo")]
    vs += tarjetas(p, 0, ["N° OC", "OC en proceso sin factura", "OC atrasadas", "Importe pendiente PEN",
                          "Importe pendiente USD", "Facturas vencidas", "Req hasta envío de OC"])
    vs += segmentadores(p, 10, [("Ordenes", "Comprador"), ("Ordenes", "Moneda"), ("Ordenes", "Mes OC")])
    a, b, c = fila(3)
    vs.append(visual(p, 20, "donutChart", *a, {"Category": [col("Ordenes", "Categoría")], "Y": [med("N° OC")]},
                     titulo="Órdenes por categoría",
                     objetos={**colores("Ordenes", "Categoría", COLORES_CATEGORIA),
                              "labels": [{"properties": {"labelStyle": _lit("Data value, percent of total"),
                                                         "labelDisplayUnits": _lit(1)}}]}))
    imp = med("Importe pendiente")
    vs.append(visual(p, 21, "barChart", *b, {"Category": [col("Facturas", "Proveedor")], "Y": [imp]},
                     titulo="Importe pendiente de pago por proveedor", orden=(imp, True),
                     objetos={"dataPoint": [{"properties": {"fill": _color("#2563EB")}}]}))
    vs.append(visual(p, 22, "columnChart", *c,
                     {"Category": [col("Requerimientos", "Etapa")], "Series": [col("Requerimientos", "Plazo")],
                      "Y": [med("N° requerimientos")]},
                     titulo="Requerimientos abiertos por etapa y plazo",
                     objetos=colores("Requerimientos", "Plazo", COLORES_PLAZO),
                     filtros=[filtro_in("fAbiertos", "Requerimientos", "Abierto / cerrado", ["Abierto"])]))
    vs.append(visual(p, 23, "tableEx", M, TABLA[0], ANCHO - 2 * M, TABLA[1], {"Values": [
        col("Ordenes", c) for c in ("OC", "Comprador", "Proveedor", "Fecha OC", "Días desde la OC", "Plazo OC",
                                    "Estado de seguimiento", "Responsable", "Último comentario")]},
        titulo="OC en proceso sin factura",
        filtros=[filtro_in("fSinFact", "Ordenes", "Categoría", ["En proceso sin factura"])]))
    return "Resumen", vs


def pagina_facturas():
    p = "facturas"
    vs = [encabezado(p, "Facturas", "Pendientes de pago, vencimientos y proveedores")]
    vs += tarjetas(p, 0, ["Facturas pendientes", "Facturas vencidas", "Facturas por vencer", "Importe pendiente PEN",
                          "Importe pendiente USD", "Importe vencido"],
                   acentos=["#2563EB", ROJO, AMBAR, "#0EA5A4", "#7C3AED", ROJO])
    vs += segmentadores(p, 10, [("Facturas", "Comprador"), ("Facturas", "Moneda"), ("Facturas", "Proveedor")])
    a, b = fila(2)
    imp = med("Importe pendiente")
    vs.append(visual(p, 20, "columnChart", *a,
                     {"Category": [col("Facturas", "Mes vencimiento")], "Series": [col("Facturas", "Plazo factura")],
                      "Y": [imp]}, titulo="Importe pendiente por mes de vencimiento",
                     objetos=colores("Facturas", "Plazo factura", COLORES_FACTURA)))
    venc = med("Importe vencido")
    vs.append(visual(p, 21, "barChart", *b, {"Category": [col("Facturas", "Proveedor")], "Y": [venc]},
                     titulo="Importe vencido por proveedor", orden=(venc, True),
                     objetos={"dataPoint": [{"properties": {"fill": _color(ROJO)}}]}))
    vs.append(visual(p, 22, "tableEx", M, TABLA[0], ANCHO - 2 * M, TABLA[1], {"Values": [
        col("Facturas", c) for c in ("OC", "Proveedor", "N° factura", "Comprador", "Moneda", "Importe pendiente",
                                     "Emisión", "Vencimiento", "Días para vencer", "Plazo factura")]},
        titulo="Facturas pendientes de pago",
        filtros=[filtro_in("fPend", "Facturas", "Estado factura", ["Pendiente de pago"]),
                 filtro_in("fVig", "Facturas", "Situación", ["Vigente"])]))
    return "Facturas", vs


def pagina_requerimientos():
    p = "requerimientos"
    vs = [encabezado(p, "Requerimientos", "Avance de cada solicitud hasta el envío de la OC y el pago")]
    vs += tarjetas(p, 0, ["Req abiertos", "Req hasta envío de OC", "Req atrasados", "Req en atención",
                          "Días promedio en la etapa"], acentos=["#2563EB", "#7C3AED", ROJO, AMBAR, "#0EA5A4"])
    vs += segmentadores(p, 10, [("Requerimientos", "Comprador"), ("Requerimientos", "Tipo"),
                                ("Requerimientos", "Controller")])
    abiertos = [filtro_in("fAbiertos", "Requerimientos", "Abierto / cerrado", ["Abierto"])]
    a, b = fila(2)
    vs.append(visual(p, 20, "columnChart", *a,
                     {"Category": [col("Requerimientos", "Etapa")], "Series": [col("Requerimientos", "Plazo")],
                      "Y": [med("N° requerimientos")]}, titulo="Abiertos por etapa y plazo", filtros=abiertos,
                     objetos=colores("Requerimientos", "Plazo", COLORES_PLAZO)))
    n = med("Req abiertos")
    vs.append(visual(p, 21, "barChart", *b, {"Category": [col("Requerimientos", "Controller")], "Y": [n]},
                     titulo="Abiertos por controller", orden=(n, True),
                     objetos={"dataPoint": [{"properties": {"fill": _color("#7C3AED")}}]}))
    vs.append(visual(p, 22, "tableEx", M, TABLA[0], ANCHO - 2 * M, TABLA[1], {"Values": [
        col("Requerimientos", c) for c in ("COD", "Comprador", "Proyecto", "Referencia", "Proveedor", "Etapa",
                                           "Días en la etapa", "Plazo", "Orden", "Controller")]},
        titulo="Requerimientos abiertos", filtros=list(abiertos)))
    return "Requerimientos", vs


def pagina_tiempos():
    p = "tiempos"
    vs = [encabezado(p, "Tiempos del flujo", "Días promedio de cada tramo (los anticipos no cuentan)")]
    tramos = ["Días solicitud → envío de OC", "Días envío de OC → factura", "Días solicitud → pago"]
    colores_t = {"Días solicitud → envío de OC": "#7C3AED", "Días envío de OC → factura": "#0EA5A4",
                 "Días solicitud → pago": "#2563EB"}
    vs += tarjetas(p, 0, tramos, acentos=list(colores_t.values()), tamano=22, etiqueta=10)
    vs += segmentadores(p, 10, [("Requerimientos", "Comprador"), ("Requerimientos", "Tipo"),
                                ("Requerimientos", "Controller")])
    a, b = fila(2)
    vs.append(visual(p, 20, "clusteredColumnChart", *a,
                     {"Category": [col("Requerimientos", "Tipo")], "Y": [med(m) for m in tramos]},
                     titulo="Días promedio por tipo de requerimiento", objetos=colores_medidas(colores_t)))
    vs.append(visual(p, 21, "lineChart", *b,
                     {"Category": [col("Requerimientos", "Mes solicitud")], "Y": [med(tramos[0])]},
                     titulo="Días solicitud → envío de OC, por mes de solicitud",
                     objetos=colores_medidas({tramos[0]: "#7C3AED"})))
    vs.append(visual(p, 22, "tableEx", M, TABLA[0], ANCHO - 2 * M, TABLA[1], {"Values": [
        col("Requerimientos", "Controller"), med("N° requerimientos")] + [med(m) for m in tramos]},
        titulo="Tiempos por controller (días promedio)"))
    return "Tiempos", vs


PAGINAS = [pagina_kpis, pagina_resumen, pagina_facturas, pagina_requerimientos, pagina_tiempos]


# ---------------------------------------------------------------- tema
def aplicar_tema():
    """Registra tema_seguimiento.json como tema personalizado del informe."""
    os.makedirs(RECURSOS, exist_ok=True)
    shutil.copyfile(TEMA, os.path.join(RECURSOS, NOMBRE_TEMA))
    ruta = os.path.join(REPORTE, "report.json")
    rep = json.load(open(ruta, encoding="utf-8"))
    base = rep["themeCollection"]["baseTheme"]
    rep["themeCollection"]["customTheme"] = {"name": NOMBRE_TEMA, "reportVersionAtImport": base["reportVersionAtImport"],
                                             "type": "RegisteredResources"}
    paquetes = [p for p in rep.get("resourcePackages", []) if p["name"] != "RegisteredResources"]
    paquetes.append({"name": "RegisteredResources", "type": "RegisteredResources",
                     "items": [{"name": NOMBRE_TEMA, "path": NOMBRE_TEMA, "type": "CustomTheme"}]})
    rep["resourcePackages"] = paquetes
    json.dump(rep, open(ruta, "w", encoding="utf-8"), ensure_ascii=False, indent=2)


def main():
    paginas = [(_id("pagina", f.__name__), *f()) for f in PAGINAS]
    modelo = leer_modelo()
    faltan = sorted({(t, c) for t, c, es_med in USADOS
                     if t not in modelo or c not in modelo[t]["medidas" if es_med else "columnas"]})
    if faltan:
        print("No se escribió nada: faltan en el modelo (¿ejecutaste actualizar_modelo.py?):")
        for t, c in faltan:
            print(f"  {t}[{c}]")
        sys.exit(1)
    aplicar_tema()
    dir_pag = os.path.join(REPORTE, "pages")
    meta = json.load(open(os.path.join(dir_pag, "pages.json"), encoding="utf-8"))
    nuevas = [pid for pid, _, _ in paginas]
    for pid, titulo, vs in paginas:
        carpeta = os.path.join(dir_pag, pid)
        # Se sobrescribe en el lugar (OneDrive bloquea un rato las carpetas recién sincronizadas);
        # los visuales que ya no existen en esta versión se borran
        os.makedirs(os.path.join(carpeta, "visuals"), exist_ok=True)
        actuales = {v["name"] for v in vs}
        for viejo in os.listdir(os.path.join(carpeta, "visuals")):
            if viejo not in actuales:
                shutil.rmtree(os.path.join(carpeta, "visuals", viejo), ignore_errors=True)
        json.dump({"$schema": f"{ESQ}/page/2.1.0/schema.json", "name": pid, "displayName": titulo,
                   "displayOption": "FitToPage", "height": ALTO, "width": ANCHO},
                  open(os.path.join(carpeta, "page.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=2)
        for v in vs:
            os.makedirs(os.path.join(carpeta, "visuals", v["name"]), exist_ok=True)
            json.dump(v, open(os.path.join(carpeta, "visuals", v["name"], "visual.json"), "w", encoding="utf-8"),
                      ensure_ascii=False, indent=2)
        print(f"Página «{titulo}»: {len(vs)} visuales")
    meta["pageOrder"] = nuevas + [p for p in meta["pageOrder"] if p not in nuevas]
    meta["activePageName"] = nuevas[0]
    json.dump(meta, open(os.path.join(dir_pag, "pages.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=2)
    print("Listo. Abre Seguimiento.pbip en Power BI Desktop.")


if __name__ == "__main__":
    main()
