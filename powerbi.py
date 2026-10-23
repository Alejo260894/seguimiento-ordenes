"""
Datos para Power BI: «Datos Power BI.xlsx», con una tabla de Excel por tema, lista para conectar.

Tablas (se relacionan por el N° de OC):
- Ordenes: una fila por OC (todas las que conoce la base) con su categoría, status, facturación y seguimiento.
- Facturas: una fila por factura, con días para vencer y plazo.
- Requerimientos: una fila por requerimiento, con etapa, plazo, fechas del flujo y días de cada tramo.
- Comentarios: historial de cada OC.
- Proveedores: una fila por RUC con su puntaje de desempeño (scorecard).
- Lineas: líneas de cada OC (reporte de detalle) con proyecto, unidad de negocio, área e ingreso de material.
- Info: fecha del reporte SAP y de la exportación.

KPI de abastecimiento (definiciones):
- Ahorro = monto referencial − monto de la OC (requerimientos con ambos datos).
- OTIF: a tiempo (conformidad ≤ fecha fin) y completo (OC facturada al 100 %).
- Lead time = fecha de conformidad − fecha de solicitud.
- Cumplimiento: OC cerrada facturada por el monto aprobado (±0.5 %).
- Desempeño del proveedor = Σ (criterio × peso) con PESOS, sobre los criterios con datos suficientes.
La app lo regenera en cada carga del reporte SAP, al cerrarse y con el botón de la pestaña Indicadores.
En Power BI basta con «Actualizar».
"""
import os
from datetime import date, datetime

import consulta as cq
import requerimientos as rq
import semaforo as sem
import seguimiento as sgm
from detalle_sap import SIN_DETALLE

ARCHIVO = "Datos Power BI.xlsx"
MON = {"SOL": "PEN"}             # SAP usa SOL en facturas y PEN en la OC
TOLERANCIA = 0.005               # facturado vs OC: ±0.5 % se considera exacto
DIAS_FACTURA = 30                # puntualidad: primera factura hasta 30 días después de la OC
PESOS = {"Cumplimiento": 0.4, "Puntualidad": 0.3, "OTIF": 0.3}     # scorecard de proveedores
MIN_EVAL = 3                     # mínimo de casos para que un criterio cuente en el puntaje


def cumplimiento(sap, facturado):
    """Cumplimiento de la compra según SAP: (categoría, evaluada)."""
    if sap is None:
        return "Sin datos de SAP", False
    pct = facturado / sap["total"] if sap["total"] else None
    if pct is not None and pct > 1 + TOLERANCIA:
        return "Sobrefacturada", True
    if sap["estado"] == "Cancelado":
        return "Cancelada", False
    if sap["estado"] != "Cerrado":
        return "En curso", False
    if sgm.con_nota_entrada(sap.get("serie"), sap["estado"], sap.get("fecha"), sap.get("ingresado")) and (
            pct is None or pct < 1 - TOLERANCIA):
        return "Atendida, pendiente de factura", False        # nota de entrada: la factura aún no se ingresa
    if not facturado:
        return "Cerrada sin factura", True
    if pct is not None and pct >= 1 - TOLERANCIA:
        return "Cumple", True
    return "Cerrada con saldo sin facturar", True


def clasificacion(puntaje):
    if puntaje is None:
        return "Sin datos suficientes"
    return "A (≥ 85)" if puntaje >= 85 else "B (70 a 85)" if puntaje >= 70 else "C (< 70)"


def _d(iso):
    return datetime.strptime(iso[:10], "%Y-%m-%d") if iso else None


def _dt(texto):
    return datetime.strptime(texto, "%Y-%m-%d %H:%M:%S") if texto else None


def tablas(con, umbrales):
    """{nombre de tabla: (columnas, filas)} con los datos actuales de la base."""
    hoy = date.today()
    filas_req = rq.listar(con)
    universo = cq.universo(con, filas_req)
    n_com = dict(con.execute("SELECT numero_oc, COUNT(*) FROM comentarios WHERE tipo='usuario' GROUP BY numero_oc"))
    det = {r["numero_oc"]: dict(r) for r in con.execute("SELECT * FROM detalle_oc")}     # reporte de detalle
    ult = sgm.ultimos_comentarios(con)

    ordenes, kpi_oc = [], {}
    for x in sorted(universo.values(), key=lambda x: x["oc"]):
        sap, o = x["sap"], x["orden"]
        vig = [f for f in x["facturas"] if f["situacion"] == sgm.FACT_VIGENTE]
        pend = [f for f in vig if f["estado_factura"] != "Pagado"]
        venc = [f for f in pend if f["vencimiento"] and f["vencimiento"] < hoy.isoformat()]
        facturado = sum(f["ft_sin_imp"] or 0 for f in vig)
        activa = o is not None and o["situacion"] == sgm.SIT_ACTIVA
        fecha = _d(x["fecha"])
        cumple, evaluada = cumplimiento(sap, facturado)
        emisiones = [f["emision"] for f in vig if f["emision"]]
        dias_fact = (_d(min(emisiones)) - fecha).days if emisiones and fecha else None
        dias_fact = dias_fact if dias_fact is None or dias_fact >= 0 else None    # factura antes de la OC: regularización
        kpi_oc[x["oc"]] = k = {
            "cumple": cumple, "evaluada": evaluada, "dias_fact": dias_fact,
            "completa": None if sap is None or not sap["total"] else
                        "Sí" if facturado >= sap["total"] * (1 - TOLERANCIA) else
                        "No" if sap["estado"] in ("Cerrado", "Cancelado")
                                and not sgm.con_nota_entrada(sap["serie"], sap["estado"], sap["fecha"],
                                                             sap.get("ingresado")) else None,
            "plazo": sem.texto(sem.nivel_oc(o, umbrales)) if activa else ""}
        ordenes.append([
            x["oc"], x["comprador"], x["proveedor"], x["ruc"], fecha, _d(sap["aprobacion"]) if sap else None,
            fecha.strftime("%Y-%m") if fecha else None, MON.get(x["mon"], x["mon"]), x["total"], x["estado_sap"],
            x["categoria"], x["status"], len(vig), len(pend), len(venc), facturado,
            facturado / x["total"] if x["total"] else None, sum(f["imp_pendiente"] or 0 for f in pend),
            o["situacion"] if o else "", o["estado_seg"] if activa else "", (o["responsable"] or "") if activa else "",
            _d(o["fecha_compromiso"]) if activa else None, (hoy - fecha.date()).days if fecha else None,
            sem.texto(sem.nivel_oc(o, umbrales)) if activa else "", n_com.get(x["oc"], 0),
            ult[x["oc"]][1] if x["oc"] in ult else "", _dt(ult[x["oc"]][0]) if x["oc"] in ult else None,
            ", ".join(f"{q['id']:03d}" for q in x["reqs"]), k["cumple"], "Sí" if evaluada else "No",
            facturado / x["total"] - 1 if evaluada and x["total"] else None, dias_fact,
            None if dias_fact is None else "Sí" if dias_fact <= DIAS_FACTURA else "No",
            *ciclo(x, det.get(x["oc"]), fecha, vig)])
    col_ord = ["OC", "Comprador", "Proveedor", "RUC", "Fecha OC", "Aprobación OC", "Mes OC", "Moneda", "Total sin IGV",
               "Estado OC", "Categoría", "Status", "N° facturas", "Facturas pendientes", "Facturas vencidas",
               "Facturado sin IGV", "% facturado", "Pendiente de pago", "Seguimiento", "Estado de seguimiento",
               "Responsable", "Fecha compromiso", "Días desde la OC", "Plazo OC", "N° comentarios",
               "Último comentario", "Fecha último comentario", "Requerimientos", "Cumplimiento",
               "Cumplimiento evaluado", "% desviación facturado", "Días OC → 1ra factura", "Factura a tiempo",
               "Tipo OC", "Estado logístico", "Fecha última entrada", "% atendido", "Proyecto", "Unidad de negocio",
               "Área", "Días OC → aprobación", "Días OC → entrada", "Días factura → pago"]

    facturas = []
    for f in sgm.listar_facturas(con):
        dv = (_d(f["vencimiento"]).date() - hoy).days if f["vencimiento"] else None
        facturas.append([
            f["numero_oc"], f["usuario"], f["proveedor"], f["ruc"], f["n_factura"], f["tipo"],
            f["estado_factura"] or "(sin estado)", f["situacion"], MON.get(f["mon"], f["mon"]), f["ft_sin_imp"],
            f["imp_pendiente"], _d(f["emision"]), _d(f["vencimiento"]),
            _d(f["vencimiento"]).strftime("%Y-%m") if f["vencimiento"] else None, _d(f["fecha_pago"]),
            dv if f["estado_factura"] != "Pagado" else None, sem.TEXTO_FACT.get(sem.nivel_factura(f, umbrales), ""),
            f["estado_oc"]])
    col_fac = ["OC", "Comprador", "Proveedor", "RUC", "N° factura", "Tipo", "Estado factura", "Situación", "Moneda",
               "FT sin impuestos", "Importe pendiente", "Emisión", "Vencimiento", "Mes vencimiento", "Fecha de pago",
               "Días para vencer", "Plazo factura", "Estado OC"]

    reqs = []
    for f in filas_req:
        dias, anticipos = [], 0
        for _, a, b in rq.TRAMOS:
            d = rq.dias_entre(f[a], f[b])
            anticipos += d is not None and d < 0
            dias.append(d if d is None or d >= 0 else None)     # negativos = anticipos: fuera de los promedios
        reqs.append([
            f["id"], f["comprador"], f["proyecto"], f["proyecto_nombre"], f["controller"], f["controller2"], f["jefe"],
            f["referencia"], f["tipo"], f["proveedor"], f["estatus"], f["monto"], f["moneda"], f["orden"],
            f["sap_estado"], f["etapa"], "Sí" if f["etapa"] in rq.ETAPAS_HASTA_OC else "No",
            "Cerrado" if f["etapa"] in rq.CERRADAS else "Abierto", f["dias_etapa"],
            sem.texto(sem.nivel_req(f, umbrales)), _d(f["f_solic"]).strftime("%Y-%m") if f["f_solic"] else None,
            *(_d(f[c]) for c in rq.CAMPOS_FECHA), *dias, anticipos, *kpi_req(f, kpi_oc)])
    col_req = (["COD", "Comprador", "Proyecto", "Nombre del proyecto", "Controller", "Controller 2", "Jefe de proyecto",
                "Referencia", "Tipo", "Proveedor", "Estatus", "Monto", "Moneda", "Orden", "Estado OC SAP", "Etapa",
                "Hasta envío de OC", "Abierto / cerrado", "Días en la etapa", "Plazo", "Mes solicitud"]
               + [rq.ETIQUETAS[c] for c in rq.CAMPOS_FECHA] + [f"Días: {t[0]}" for t in rq.TRAMOS]
               + ["Tramos con anticipo", "Monto referencial", "Tipo de referencia", "Ahorro", "Base del ahorro",
                  "% ahorro", "Fecha de conformidad", "A tiempo", "Completo", "OTIF", "Lead time (días)"])

    provs = proveedores(universo, kpi_oc, reqs, col_req)
    comentarios = [[r["numero_oc"], _dt(r["fecha"]), r["autor"], r["tipo"], r["texto"]]
                   for r in con.execute("SELECT * FROM comentarios ORDER BY numero_oc, fecha, id")]
    u = sgm.ultima_carga(con)
    info = [["Reporte SAP", _dt(u["fecha_reporte"]) if u else None], ["Exportado", datetime.now().replace(microsecond=0)]]
    return {
        "Ordenes": (col_ord, ordenes),
        "Facturas": (col_fac, facturas),
        "Requerimientos": (col_req, reqs),
        "Comentarios": (["OC", "Fecha", "Autor", "Tipo", "Texto"], comentarios),
        "Proveedores": provs,
        "Lineas": lineas(con),
        "Info": (["Dato", "Fecha"], info),
    }


def _dias(a, b):
    """Días entre dos fechas (datetime o ISO); None si falta alguna o el resultado es negativo (regularización)."""
    a, b = (_d(v) if isinstance(v, str) else v for v in (a, b))
    if not a or not b:
        return None
    d = (b - a).days
    return d if d >= 0 else None


def ciclo(x, d, fecha, vigentes):
    """Columnas del reporte de detalle y tiempos del ciclo OC → aprobación → entrada → factura → pago."""
    d = d or {}
    sap = x["sap"] or {}
    pagos = [f["fecha_pago"] for f in vigentes if f["fecha_pago"]]
    emisiones = [f["emision"] for f in vigentes if f["emision"]]
    todas_pagadas = vigentes and all(f["estado_factura"] == "Pagado" for f in vigentes)
    return [d.get("tipo_oc") or "", d.get("estado_logistico") or SIN_DETALLE, _d(d.get("fecha_entrada")),
            d.get("pct_ingresado"), d.get("proyecto") or "", d.get("unidad_negocio") or "", d.get("area") or "",
            _dias(fecha, sap.get("aprobacion")), _dias(fecha, d.get("fecha_entrada")),
            _dias(min(emisiones), max(pagos)) if todas_pagadas and emisiones and pagos else None]


def lineas(con):
    """Tabla Lineas: una fila por línea de OC (solo OC aprobadas, del reporte de detalle)."""
    filas = [[r["numero_oc"], r["linea"], r["tipo_oc"], r["codigo"], r["descripcion"], r["cantidad"], r["um"],
              r["total_linea"], MON.get(r["moneda"], r["moneda"]), r["proyecto"], r["unidad_negocio"], r["area"],
              r["ingresado"], r["pendiente_ingreso"], r["estado_logistico"], _d(r["fecha_entrada"]), r["facturado"],
              r["estado_linea"] or ""]
             for r in con.execute("SELECT * FROM detalle_lineas ORDER BY numero_oc, linea")]
    return (["OC", "Línea", "Tipo OC", "Código", "Descripción", "Cantidad", "UM", "Total línea sin IGV", "Moneda",
             "Proyecto", "Unidad de negocio", "Área", "Cantidad ingresada", "Pendiente de ingreso", "Estado logístico",
             "Fecha de entrada", "Facturado sin IGV", "Estado de la línea"], filas)


def kpi_req(f, kpi_oc):
    """Ahorro, OTIF y lead time de un requerimiento (columnas finales de la tabla Requerimientos)."""
    ref, monto = f["monto_ref"], f["monto"]
    ahorro = ref - monto if ref and monto is not None else None
    conf, fin = f["f_conformidad"], f["fecha_fin"]
    a_tiempo = None if not conf or not fin else "Sí" if conf <= fin else "No"
    completo = (kpi_oc.get(f["orden"]) or {}).get("completa") if conf else None
    otif = None if a_tiempo is None or completo is None else "Sí" if a_tiempo == completo == "Sí" else "No"
    if a_tiempo == "No" or completo == "No":
        otif = "No"
    lead = rq.dias_entre(f["f_solic"], conf)
    return [ref, f["tipo_ref"] or None, ahorro, ref if ahorro is not None else None,
            ahorro / ref if ahorro is not None else None, _d(conf), a_tiempo, completo, otif,
            lead if lead is not None and lead >= 0 else None]


def proveedores(universo, kpi_oc, reqs, col_req):
    """Scorecard por RUC: cumplimiento, puntualidad de facturación y OTIF, ponderados con PESOS."""
    i_orden, i_otif = col_req.index("Orden"), col_req.index("OTIF")
    otif_oc = {}
    for r in reqs:
        if r[i_orden] and r[i_otif]:
            otif_oc.setdefault(r[i_orden], []).append(r[i_otif] == "Sí")
    grupos = {}
    for x in universo.values():
        if x["ruc"]:
            grupos.setdefault(x["ruc"], []).append(x)
    filas = []
    for ruc, ocs in grupos.items():
        ocs.sort(key=lambda x: x["fecha"] or "")
        k = [kpi_oc[x["oc"]] for x in ocs]
        evaluadas = [v for v in k if v["evaluada"]]
        cumple = sum(v["cumple"] == "Cumple" for v in evaluadas)
        con_fact = [v["dias_fact"] for v in k if v["dias_fact"] is not None]
        puntuales = sum(d <= DIAS_FACTURA for d in con_fact)
        otif = [ok for x in ocs for ok in otif_oc.get(x["oc"], [])]
        criterios = {"Cumplimiento": (cumple, len(evaluadas)), "Puntualidad": (puntuales, len(con_fact)),
                     "OTIF": (sum(otif), len(otif))}
        pct = {c: (a / n if n else None) for c, (a, n) in criterios.items()}
        validos = {c: pct[c] for c, (_, n) in criterios.items() if n >= MIN_EVAL}
        peso = sum(PESOS[c] for c in validos)
        puntaje = sum(validos[c] * 100 * PESOS[c] for c in validos) / peso if peso else None
        filas.append([ruc, ocs[-1]["proveedor"], ocs[-1]["comprador"], len(ocs), _d(ocs[-1]["fecha"]),
                      len(evaluadas), pct["Cumplimiento"], len(con_fact), pct["Puntualidad"],
                      round(sum(con_fact) / len(con_fact), 1) if con_fact else None, len(otif), pct["OTIF"],
                      sum(v["plazo"] == "Atrasado" for v in k), puntaje, clasificacion(puntaje),
                      ", ".join(validos) or "-"])
    columnas = ["RUC", "Proveedor", "Último comprador", "N° OC", "Última OC", "OC evaluadas", "% cumplimiento",
                "OC con factura", "% puntualidad", "Días promedio OC → factura", "Entregas evaluadas (OTIF)", "% OTIF",
                "OC sin factura atrasadas", "Puntaje", "Clasificación", "Criterios usados"]
    return columnas, sorted(filas, key=lambda f: (f[13] is None, -(f[13] or 0)))


def exportar(con, carpeta, umbrales):
    """Escribe «Datos Power BI.xlsx» en `carpeta` (reemplaza el anterior). Devuelve la ruta.
    PermissionError si el archivo está abierto en Excel."""
    import xlsxwriter
    destino = os.path.join(carpeta, ARCHIVO)
    tmp = os.path.join(carpeta, f"~tmp_pbi_{os.getpid()}.xlsx")
    wb = xlsxwriter.Workbook(tmp, {"strings_to_formulas": False, "remove_timezone": True})
    fecha = wb.add_format({"num_format": "dd/mm/yyyy"})
    fecha_hora = wb.add_format({"num_format": "dd/mm/yyyy hh:mm"})
    numero = wb.add_format({"num_format": "#,##0.00"})
    pct = wb.add_format({"num_format": "0.0%"})
    for nombre, (columnas, filas) in tablas(con, umbrales).items():
        ws = wb.add_worksheet(nombre)
        cols = []
        for i, c in enumerate(columnas):
            valores = [f[i] for f in filas if f[i] is not None]
            es_fecha = valores and all(isinstance(v, datetime) for v in valores)
            hora = es_fecha and any(v.hour or v.minute for v in valores)
            fmt = (fecha_hora if hora else fecha) if es_fecha else pct if c.startswith("%") else \
                numero if valores and all(isinstance(v, float) for v in valores) else None
            cols.append({"header": c, **({"format": fmt} if fmt else {})})
            ws.set_column(i, i, 12 if es_fecha else 14 if fmt else 22 if c not in ("Status", "Texto", "Último comentario") else 50)
        ws.add_table(0, 0, max(len(filas), 1), len(columnas) - 1,
                     {"name": nombre, "columns": cols, "data": filas, "style": "Table Style Light 9"})
        ws.freeze_panes(1, 1)
    wb.close()
    try:
        os.replace(tmp, destino)
    except PermissionError:
        os.remove(tmp)
        raise
    return destino
