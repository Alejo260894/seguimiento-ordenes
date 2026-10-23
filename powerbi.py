"""
Datos para Power BI: «Datos Power BI.xlsx», con una tabla de Excel por tema, lista para conectar.

Tablas (se relacionan por el N° de OC):
- Ordenes: una fila por OC (todas las que conoce la base) con su categoría, status, facturación y seguimiento.
- Facturas: una fila por factura, con días para vencer y plazo.
- Requerimientos: una fila por requerimiento, con etapa, plazo, fechas del flujo y días de cada tramo.
- Comentarios: historial de cada OC.
- Info: fecha del reporte SAP y de la exportación.
La app lo regenera en cada carga del reporte SAP, al cerrarse y con el botón de la pestaña Indicadores.
En Power BI basta con «Actualizar».
"""
import os
from datetime import date, datetime

import consulta as cq
import requerimientos as rq
import semaforo as sem
import seguimiento as sgm

ARCHIVO = "Datos Power BI.xlsx"
MON = {"SOL": "PEN"}             # SAP usa SOL en facturas y PEN en la OC


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
    ult = sgm.ultimos_comentarios(con)

    ordenes = []
    for x in sorted(universo.values(), key=lambda x: x["oc"]):
        sap, o = x["sap"], x["orden"]
        vig = [f for f in x["facturas"] if f["situacion"] == sgm.FACT_VIGENTE]
        pend = [f for f in vig if f["estado_factura"] != "Pagado"]
        venc = [f for f in pend if f["vencimiento"] and f["vencimiento"] < hoy.isoformat()]
        facturado = sum(f["ft_sin_imp"] or 0 for f in vig)
        activa = o is not None and o["situacion"] == sgm.SIT_ACTIVA
        fecha = _d(x["fecha"])
        ordenes.append([
            x["oc"], x["comprador"], x["proveedor"], x["ruc"], fecha, _d(sap["aprobacion"]) if sap else None,
            fecha.strftime("%Y-%m") if fecha else None, MON.get(x["mon"], x["mon"]), x["total"], x["estado_sap"],
            x["categoria"], x["status"], len(vig), len(pend), len(venc), facturado,
            facturado / x["total"] if x["total"] else None, sum(f["imp_pendiente"] or 0 for f in pend),
            o["situacion"] if o else "", o["estado_seg"] if activa else "", (o["responsable"] or "") if activa else "",
            _d(o["fecha_compromiso"]) if activa else None, (hoy - fecha.date()).days if fecha else None,
            sem.texto(sem.nivel_oc(o, umbrales)) if activa else "", n_com.get(x["oc"], 0),
            ult[x["oc"]][1] if x["oc"] in ult else "", _dt(ult[x["oc"]][0]) if x["oc"] in ult else None,
            ", ".join(f"{q['id']:03d}" for q in x["reqs"])])
    col_ord = ["OC", "Comprador", "Proveedor", "RUC", "Fecha OC", "Aprobación OC", "Mes OC", "Moneda", "Total sin IGV",
               "Estado OC", "Categoría", "Status", "N° facturas", "Facturas pendientes", "Facturas vencidas",
               "Facturado sin IGV", "% facturado", "Pendiente de pago", "Seguimiento", "Estado de seguimiento",
               "Responsable", "Fecha compromiso", "Días desde la OC", "Plazo OC", "N° comentarios",
               "Último comentario", "Fecha último comentario", "Requerimientos"]

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
            *(_d(f[c]) for c in rq.CAMPOS_FECHA), *dias, anticipos])
    col_req = (["COD", "Comprador", "Proyecto", "Nombre del proyecto", "Controller", "Controller 2", "Jefe de proyecto",
                "Referencia", "Tipo", "Proveedor", "Estatus", "Monto", "Moneda", "Orden", "Estado OC SAP", "Etapa",
                "Hasta envío de OC", "Abierto / cerrado", "Días en la etapa", "Plazo", "Mes solicitud"]
               + [rq.ETIQUETAS[c] for c in rq.CAMPOS_FECHA] + [f"Días: {t[0]}" for t in rq.TRAMOS]
               + ["Tramos con anticipo"])

    comentarios = [[r["numero_oc"], _dt(r["fecha"]), r["autor"], r["tipo"], r["texto"]]
                   for r in con.execute("SELECT * FROM comentarios ORDER BY numero_oc, fecha, id")]
    u = sgm.ultima_carga(con)
    info = [["Reporte SAP", _dt(u["fecha_reporte"]) if u else None], ["Exportado", datetime.now().replace(microsecond=0)]]
    return {
        "Ordenes": (col_ord, ordenes),
        "Facturas": (col_fac, facturas),
        "Requerimientos": (col_req, reqs),
        "Comentarios": (["OC", "Fecha", "Autor", "Tipo", "Texto"], comentarios),
        "Info": (["Dato", "Fecha"], info),
    }


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
