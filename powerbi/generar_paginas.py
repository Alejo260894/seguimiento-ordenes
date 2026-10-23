"""
Genera las páginas del informe de Power BI (proyecto Seguimiento.pbip, formato PBIR).

    python generar_paginas.py

- Crea o reemplaza las páginas Resumen, Facturas, Requerimientos y Tiempos. Las páginas creadas a mano
  en Power BI no se tocan.
- Antes de escribir, comprueba que cada tabla, columna y medida exista en el modelo
  (Seguimiento.SemanticModel). Si falta alguna, no escribe nada.
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
MODELO = os.path.join(AQUI, "Seguimiento.SemanticModel", "definition", "tables")
ESQ = "https://developer.microsoft.com/json-schemas/fabric/item/report/definition"
ANCHO, ALTO = 1280, 720


def _id(*partes):
    return hashlib.md5("|".join(partes).encode("utf-8")).hexdigest()[:20]


def _lit(v):
    if isinstance(v, bool):
        return {"expr": {"Literal": {"Value": "true" if v else "false"}}}
    if isinstance(v, (int, float)):
        return {"expr": {"Literal": {"Value": f"{v}D"}}}
    return {"expr": {"Literal": {"Value": "'" + str(v).replace("'", "''") + "'"}}}


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


ROJO, AMBAR, VERDE, GRIS = "#D64550", "#F2B134", "#4E9F3D", "#A6A6A6"
COLORES_PLAZO = {"Atrasado": ROJO, "Atención": AMBAR, "En plazo": VERDE}
COLORES_FACTURA = {"Vencida": ROJO, "Por vencer": AMBAR, "Al día": VERDE}
COLORES_CATEGORIA = {"Pagada": VERDE, "Pendiente de pago": AMBAR, "En proceso sin factura": "#E07B39",
                     "Pago parcial": "#7BAFD4", "Cerrada sin factura": GRIS, "Cancelada sin factura": "#6E6E6E",
                     "No está en el reporte": "#D0D0D0"}


def colores(tabla, campo, por_valor):
    """objects.dataPoint: un color fijo para cada valor de tabla[campo] (series o categorías)."""
    return {"dataPoint": [
        {"properties": {"fill": {"solid": {"color": _lit(color)}}},
         "selector": {"data": [{"scopeId": {"Comparison": {
             "ComparisonKind": 0,
             "Left": {"Column": {"Expression": {"SourceRef": {"Entity": tabla}}, "Property": campo}},
             "Right": {"Literal": {"Value": "'" + valor + "'"}}}}}]}}
        for valor, color in por_valor.items()]}


# ---------------------------------------------------------------- visuales
def visual(pagina, i, tipo, x, y, w, h, roles, titulo=None, objetos=None, filtros=None, orden=None):
    nombre = _id(pagina, str(i))
    proyecciones = {}
    for rol, campos in roles.items():
        ps = []
        for k, c in enumerate(campos):
            p = dict(c)
            if k == 0 and rol in ("Category", "Values") and tipo != "tableEx":
                p["active"] = True
            ps.append(p)
        proyecciones[rol] = {"projections": ps}
    v = {"visualType": tipo, "query": {"queryState": proyecciones}, "drillFilterOtherVisuals": True}
    if orden:
        campo, desc = orden
        v["query"]["sortDefinition"] = {"sort": [{"field": campo["field"], "direction": "Descending" if desc else "Ascending"}],
                                        "isDefaultSort": True}
    if objetos:
        v["objects"] = objetos
    if titulo:
        v["visualContainerObjects"] = {"title": [{"properties": {"show": _lit(True), "text": _lit(titulo)}}]}
    cont = {"$schema": f"{ESQ}/visualContainer/2.14.0/schema.json", "name": nombre,
            "position": {"x": x, "y": y, "z": i, "height": h, "width": w, "tabOrder": i}, "visual": v}
    if filtros:
        cont["filterConfig"] = {"filters": filtros}
    return cont


def tarjetas(pagina, inicio, medidas, y=8, h=100):
    w = (ANCHO - 16 - 8 * (len(medidas) - 1)) / len(medidas)
    return [visual(pagina, inicio + k, "card", 8 + k * (w + 8), y, w, h, {"Values": [med(m)]},
                   objetos={"categoryLabels": [{"properties": {"fontSize": _lit(9)}}]})
            for k, m in enumerate(medidas)]


def segmentadores(pagina, inicio, campos, y=116, extra_medida=None):
    vs = []
    w = 230
    for k, (tabla, campo) in enumerate(campos):
        vs.append(visual(pagina, inicio + k, "slicer", 8 + k * (w + 8), y, w, 64, {"Values": [col(tabla, campo)]},
                         objetos={"data": [{"properties": {"mode": _lit("Dropdown")}}]}))
    if extra_medida:
        vs.append(visual(pagina, inicio + len(campos), "card", ANCHO - 8 - 300, y, 300, 64,
                         {"Values": [med(extra_medida)]},
                         objetos={"labels": [{"properties": {"fontSize": _lit(12)}}],
                                  "categoryLabels": [{"properties": {"show": _lit(False)}}]}))
    return vs


# ---------------------------------------------------------------- páginas
def pagina_resumen():
    p = "resumen"
    vs = tarjetas(p, 0, ["N° OC", "OC en proceso sin factura", "OC atrasadas", "Importe pendiente PEN",
                         "Importe pendiente USD", "Facturas vencidas", "Req hasta envío de OC"])
    vs += segmentadores(p, 10, [("Ordenes", "Comprador"), ("Ordenes", "Moneda"), ("Ordenes", "Mes OC")], extra_medida="Datos al")
    vs.append(visual(p, 20, "donutChart", 8, 188, 400, 280, {"Category": [col("Ordenes", "Categoría")], "Y": [med("N° OC")]},
                     titulo="Órdenes por categoría",
                     objetos={**colores("Ordenes", "Categoría", COLORES_CATEGORIA),
                              "labels": [{"properties": {"labelStyle": _lit("Data value, percent of total"),
                                                         "labelDisplayUnits": _lit(1)}}]}))
    imp = med("Importe pendiente")
    vs.append(visual(p, 21, "barChart", 416, 188, 420, 280, {"Category": [col("Facturas", "Proveedor")], "Y": [imp]},
                     titulo="Importe pendiente de pago por proveedor", orden=(imp, True)))
    vs.append(visual(p, 22, "columnChart", 844, 188, 428, 280,
                     {"Category": [col("Requerimientos", "Etapa")], "Series": [col("Requerimientos", "Plazo")],
                      "Y": [med("N° requerimientos")]},
                     titulo="Requerimientos abiertos por etapa y plazo",
                     objetos=colores("Requerimientos", "Plazo", COLORES_PLAZO),
                     filtros=[filtro_in("fAbiertos", "Requerimientos", "Abierto / cerrado", ["Abierto"])]))
    vs.append(visual(p, 23, "tableEx", 8, 476, ANCHO - 16, 236, {"Values": [
        col("Ordenes", c) for c in ("OC", "Comprador", "Proveedor", "Fecha OC", "Días desde la OC", "Plazo OC",
                                    "Estado de seguimiento", "Responsable", "Último comentario")]},
        titulo="OC en proceso sin factura",
        filtros=[filtro_in("fSinFact", "Ordenes", "Categoría", ["En proceso sin factura"])]))
    return "Resumen", vs


def pagina_facturas():
    p = "facturas"
    vs = tarjetas(p, 0, ["Facturas pendientes", "Facturas vencidas", "Facturas por vencer", "Importe pendiente PEN",
                         "Importe pendiente USD", "Importe vencido"])
    vs += segmentadores(p, 10, [("Facturas", "Comprador"), ("Facturas", "Moneda"), ("Facturas", "Proveedor")],
                        extra_medida="Datos al")
    imp = med("Importe pendiente")
    vs.append(visual(p, 20, "columnChart", 8, 188, 630, 280,
                     {"Category": [col("Facturas", "Mes vencimiento")], "Series": [col("Facturas", "Plazo factura")],
                      "Y": [imp]}, titulo="Importe pendiente por mes de vencimiento",
                     objetos=colores("Facturas", "Plazo factura", COLORES_FACTURA)))
    venc = med("Importe vencido")
    vs.append(visual(p, 21, "barChart", 646, 188, 626, 280, {"Category": [col("Facturas", "Proveedor")], "Y": [venc]},
                     titulo="Importe vencido por proveedor", orden=(venc, True)))
    vs.append(visual(p, 22, "tableEx", 8, 476, ANCHO - 16, 236, {"Values": [
        col("Facturas", c) for c in ("OC", "Proveedor", "N° factura", "Comprador", "Moneda", "Importe pendiente",
                                     "Emisión", "Vencimiento", "Días para vencer", "Plazo factura")]},
        titulo="Facturas pendientes de pago",
        filtros=[filtro_in("fPend", "Facturas", "Estado factura", ["Pendiente de pago"]),
                 filtro_in("fVig", "Facturas", "Situación", ["Vigente"])]))
    return "Facturas", vs


def pagina_requerimientos():
    p = "requerimientos"
    vs = tarjetas(p, 0, ["Req abiertos", "Req hasta envío de OC", "Req atrasados", "Req en atención",
                         "Días promedio en la etapa"])
    vs += segmentadores(p, 10, [("Requerimientos", "Comprador"), ("Requerimientos", "Tipo"),
                                ("Requerimientos", "Controller")], extra_medida="Datos al")
    abiertos = [filtro_in("fAbiertos", "Requerimientos", "Abierto / cerrado", ["Abierto"])]
    vs.append(visual(p, 20, "columnChart", 8, 188, 630, 280,
                     {"Category": [col("Requerimientos", "Etapa")], "Series": [col("Requerimientos", "Plazo")],
                      "Y": [med("N° requerimientos")]}, titulo="Abiertos por etapa y plazo", filtros=abiertos,
                     objetos=colores("Requerimientos", "Plazo", COLORES_PLAZO)))
    n = med("Req abiertos")
    vs.append(visual(p, 21, "barChart", 646, 188, 626, 280, {"Category": [col("Requerimientos", "Controller")], "Y": [n]},
                     titulo="Abiertos por controller", orden=(n, True)))
    vs.append(visual(p, 22, "tableEx", 8, 476, ANCHO - 16, 236, {"Values": [
        col("Requerimientos", c) for c in ("COD", "Comprador", "Proyecto", "Referencia", "Proveedor", "Etapa",
                                           "Días en la etapa", "Plazo", "Orden", "Controller")]},
        titulo="Requerimientos abiertos", filtros=list(abiertos)))
    return "Requerimientos", vs


def pagina_tiempos():
    p = "tiempos"
    tramos = ["Días solicitud → envío de OC", "Días envío de OC → factura", "Días solicitud → pago"]
    vs = tarjetas(p, 0, tramos)
    vs += segmentadores(p, 10, [("Requerimientos", "Comprador"), ("Requerimientos", "Tipo"),
                                ("Requerimientos", "Controller")], extra_medida="Datos al")
    vs.append(visual(p, 20, "clusteredColumnChart", 8, 188, 630, 280,
                     {"Category": [col("Requerimientos", "Tipo")], "Y": [med(m) for m in tramos]},
                     titulo="Días promedio por tipo de requerimiento"))
    vs.append(visual(p, 21, "lineChart", 646, 188, 626, 280,
                     {"Category": [col("Requerimientos", "Mes solicitud")], "Y": [med(tramos[0])]},
                     titulo="Días solicitud → envío de OC, por mes de solicitud"))
    vs.append(visual(p, 22, "tableEx", 8, 476, ANCHO - 16, 236, {"Values": [
        col("Requerimientos", "Controller"), med("N° requerimientos")] + [med(m) for m in tramos]},
        titulo="Tiempos por controller (días promedio; los anticipos no cuentan)"))
    return "Tiempos", vs


PAGINAS = [pagina_resumen, pagina_facturas, pagina_requerimientos, pagina_tiempos]


def main():
    paginas = [(_id("pagina", f.__name__), *f()) for f in PAGINAS]
    modelo = leer_modelo()
    faltan = sorted({(t, c) for t, c, es_med in USADOS
                     if t not in modelo or c not in modelo[t]["medidas" if es_med else "columnas"]})
    if faltan:
        print("No se escribió nada: faltan en el modelo:")
        for t, c in faltan:
            print(f"  {t}[{c}]")
        sys.exit(1)
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
