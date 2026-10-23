"""
Consulta de cualquier orden: reúne en un solo lugar lo que está repartido en la base.

- Último reporte SAP (sap_oc): comprador, fechas, proveedor, total y estado de la OC.
- Facturas (tabla facturas): estado, importes, vencimiento y pago.
- Seguimiento de OC sin factura (tabla ordenes) y comentarios.
- Requerimientos vinculados (req.orden).
Con eso se arma un status corto (para la lista) y uno completo para copiar y enviar.
"""
from datetime import date, datetime

import seguimiento as sgm


# Status corto -> categoría (para agrupar en gráficos: Power BI)
CATEGORIAS = [("Pagada", "Pagada"), ("Pago parcial", "Pago parcial"), ("Facturada, pendiente", "Pendiente de pago"),
              ("En proceso, sin factura", "En proceso sin factura"), ("Sin factura · cerrada", "Cerrada sin factura"),
              ("Cerrada sin factura", "Cerrada sin factura"), ("Cancelada sin factura", "Cancelada sin factura"),
              ("No está", "No está en el reporte")]


def _fmt(iso):
    return f"{iso[8:10]}/{iso[5:7]}/{iso[:4]}" if iso else ""


def _monto(mon, v):
    mon = {"SOL": "PEN"}.get(mon, mon)          # las facturas dicen SOL y la OC PEN: se muestra igual
    return f"{mon or ''} {v or 0:,.2f}".strip()


def universo(con, filas_req):
    """{N° OC: datos} de todas las órdenes que conoce la base (reporte SAP, seguimiento, facturas y requerimientos).
    `filas_req`: salida de requerimientos.listar (trae la etapa de cada requerimiento)."""
    ocs = {}

    def oc(n):
        return ocs.setdefault(int(n), {"oc": int(n), "sap": None, "orden": None, "facturas": [], "reqs": []})

    for r in con.execute("SELECT * FROM sap_oc"):
        oc(r["numero_oc"])["sap"] = dict(r)
    for r in sgm.listar(con):
        oc(r["numero_oc"])["orden"] = r
    for r in con.execute("SELECT * FROM facturas ORDER BY emision, n_factura"):
        if r["numero_oc"] is not None:
            oc(r["numero_oc"])["facturas"].append(dict(r))
    for f in filas_req:
        if f["orden"]:
            oc(f["orden"])["reqs"].append(f)
    for x in ocs.values():
        x.update(_resumen(x))
    return ocs


def _resumen(x):
    """Campos para la lista (comprador, proveedor…) y el status corto."""
    sap, o, fx = x["sap"], x["orden"], x["facturas"]
    base = sap or o or (fx[0] if fx else None)
    r = x["reqs"][0] if x["reqs"] else {}
    d = {
        "comprador": (sap or {}).get("usuario") or (o or {}).get("usuario") or (fx[0]["usuario"] if fx else "")
                     or r.get("comprador", ""),
        "proveedor": (base or {}).get("proveedor") or r.get("proveedor", ""),
        "ruc": (base or {}).get("ruc") or "",
        "fecha": (sap or {}).get("fecha") or (o or {}).get("fecha_oc") or (fx[0]["fecha_oc"] if fx else "") or "",
        "mon": (sap or {}).get("mon") or (o or {}).get("mon") or (fx[0]["mon"] if fx else "") or r.get("moneda", ""),
        "total": (sap or {}).get("total") if sap else (o or {}).get("total"),
        "estado_sap": sap["estado"] if sap else "No está en el último reporte",
    }
    d["status"], d["nivel"] = status(x)
    d["categoria"] = next((c for p, c in CATEGORIAS if d["status"].startswith(p)), "Otro")
    d["texto_busqueda"] = " ".join(str(v or "") for v in (
        x["oc"], d["comprador"], d["proveedor"], d["ruc"], d["status"],
        *(f["n_factura"] for f in fx),
        *(f"{q['id']:03d} {q['referencia']} {q['proyecto']} {q.get('proyecto_nombre', '')}" for q in x["reqs"]))).lower()
    return d


def status(x, hoy=None):
    """(status corto, nivel) — nivel: 'ok', 'pendiente', 'atencion' o 'cerrada' (para colorear la lista)."""
    hoy = (hoy or date.today()).isoformat()
    sap, o = x["sap"], x["orden"]
    vigentes = [f for f in x["facturas"] if f["situacion"] == sgm.FACT_VIGENTE]
    if sap is None:
        if o is not None:
            return f"No está en el último reporte SAP (seguimiento: {o['situacion']})", "cerrada"
        if vigentes or x["facturas"]:
            return "No está en el último reporte SAP", "cerrada"
        return "No está en el reporte SAP (solo registrada en un requerimiento)", "atencion"
    estado = sap["estado"]
    if not vigentes:
        if estado == "Cancelado":
            return "Cancelada sin factura", "cerrada"
        if estado == "Cerrado":
            return "Cerrada sin factura (no se va a facturar)", "cerrada"
        if o is not None and o["situacion"] == sgm.SIT_SIN_FACT:
            return "Sin factura · cerrada en el seguimiento (no se facturará)", "cerrada"
        seg = f" · {o['estado_seg']}" if o is not None and o["situacion"] == sgm.SIT_ACTIVA else ""
        return f"En proceso, sin factura{seg}", "atencion"
    pendientes = [f for f in vigentes if f["estado_factura"] != "Pagado"]
    facturado = sum(f["ft_sin_imp"] or 0 for f in vigentes)
    pct = facturado / sap["total"] * 100 if sap["total"] else None
    if not pendientes:
        pagos = [f["fecha_pago"] for f in vigentes if f["fecha_pago"]]
        if estado == "Cerrado" or (pct is not None and pct >= 99.5):
            return f"Pagada{' el ' + _fmt(max(pagos)) if pagos else ''}", "ok"
        return f"Pago parcial: facturas pagadas, OC facturada al {pct:.0f} %" if pct is not None else "Pago parcial", "pendiente"
    vencidas = [f for f in pendientes if f["vencimiento"] and f["vencimiento"] < hoy]
    txt = f"Facturada, pendiente de pago ({len(pendientes)} de {len(vigentes)} factura{'s' if len(vigentes) > 1 else ''}"
    txt += f", {len(vencidas)} vencida{'s' if len(vencidas) > 1 else ''})" if vencidas else ")"
    if estado == "En proceso" and pct is not None and pct < 99.5:
        txt += f" · OC facturada al {pct:.0f} %"
    return txt, "atencion" if vencidas else "pendiente"


def texto_status(x, fecha_reporte=None, comentarios=(), hoy=None):
    """Status completo de la orden, en texto para copiar (correo, chat…)."""
    hoy = hoy or date.today()
    sap, o = x["sap"], x["orden"]
    L = [f"OC {x['oc']} · {x['proveedor'] or '(sin proveedor)'}" + (f" (RUC {x['ruc']})" if x["ruc"] else "")]
    if sap:
        L.append(f"Comprador: {sap['usuario']} · Fecha OC: {_fmt(sap['fecha'])}"
                 + (f" · Aprobación: {_fmt(sap['aprobacion'])}" if sap["aprobacion"] else ""))
        L.append(f"Total sin IGV: {_monto(sap['mon'], sap['total'])} · Estado en SAP: {sap['estado']}")
    elif x["comprador"]:
        L.append(f"Comprador: {x['comprador']}")
    L += [f"STATUS: {x['status']}", ""]

    vigentes = [f for f in x["facturas"] if f["situacion"] == sgm.FACT_VIGENTE]
    otras = [f for f in x["facturas"] if f["situacion"] != sgm.FACT_VIGENTE]
    if vigentes:
        L.append(f"Facturas ({len(vigentes)}):")
        for f in vigentes:
            linea = f"  • {f['n_factura']} · {f['estado_factura'] or 'sin estado'} · {_monto(f['mon'], f['ft_sin_imp'])} sin imp."
            if f["emision"]:
                linea += f" · emitida {_fmt(f['emision'])}"
            if f["estado_factura"] == "Pagado":
                linea += f" · pagada {_fmt(f['fecha_pago'])}" if f["fecha_pago"] else ""
            elif f["vencimiento"]:
                dias = (datetime.strptime(f["vencimiento"], "%Y-%m-%d").date() - hoy).days
                cuando = f"vencida hace {-dias} días" if dias < 0 else "vence hoy" if dias == 0 else f"en {dias} días"
                linea += f" · vence {_fmt(f['vencimiento'])} ({cuando})"
                if f["imp_pendiente"]:
                    linea += f" · pendiente {_monto(f['mon'], f['imp_pendiente'])}"
            L.append(linea)
        if sap and sap["total"]:
            facturado = sum(f["ft_sin_imp"] or 0 for f in vigentes)
            L.append(f"  Facturado: {_monto(sap['mon'], facturado)} de {_monto(sap['mon'], sap['total'])} "
                     f"({facturado / sap['total'] * 100:.0f} %)")
    else:
        L.append("Facturas: ninguna")
    if otras:
        L.append("  Ya no aparecen en el reporte (¿anuladas?): " + ", ".join(f["n_factura"] for f in otras))

    L.append("")
    if x["reqs"]:
        L.append("Requerimientos:")
        for q in x["reqs"]:
            det = " · ".join(v for v in (q["referencia"], f"proyecto {q['proyecto']}" if q["proyecto"] else "",
                                          q.get("proyecto_nombre", "")) if v)
            L.append(f"  • COD {q['id']:03d} · {det} · etapa: {q['etapa']}"
                     + (f" · {q['comprador']}" if q.get("comprador") and sap and q["comprador"] != sap["usuario"] else ""))
    else:
        L.append("Requerimientos: ninguno vinculado a esta OC")

    if o is not None:
        seg = f"Seguimiento de OC sin factura: {o['situacion']}"
        if o["situacion"] == sgm.SIT_ACTIVA:
            seg += f" · {o['estado_seg']}"
            if o["responsable"]:
                seg += f" · responsable {o['responsable']}"
            if o["fecha_compromiso"]:
                seg += f" · compromiso {_fmt(o['fecha_compromiso'])}"
        L.append(seg)
    usuario = [c for c in comentarios if c["tipo"] == "usuario"]
    if usuario:
        c = usuario[-1]
        L.append(f"Último comentario ({_fmt(c['fecha'][:10])}, {c['autor']}): {c['texto']}")
    if fecha_reporte:
        L += ["", f"Según el reporte SAP del {_fmt(fecha_reporte[:10])} {fecha_reporte[11:16]}."]
    return "\n".join(L)


def donde(x):
    """En qué pestañas de la app aparece la orden: {'oc': bool, 'fac': bool, 'req': [ids]}."""
    return {"oc": x["orden"] is not None, "fac": bool(x["facturas"]), "req": [q["id"] for q in x["reqs"]]}
