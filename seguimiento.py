"""
Base de seguimiento de órdenes sin factura (SQLite).

Cada carga de un reporte SAP sincroniza la base:
- OC sin ninguna factura vigente y que NO está Cerrada ni Cancelada -> entra (o sigue) en seguimiento.
  (Una OC Cerrada o Cancelada sin factura ya no se va a facturar: no se trata.)
- OC en seguimiento que ya tiene factura -> se cierra como "Facturada" (queda el historial).
- OC en seguimiento que pasa a Cerrada/Cancelada sin factura -> se cierra como "OC cerrada/cancelada sin factura".
- OC en seguimiento que ya no aparece en el reporte -> se cierra como "Fuera del reporte".
Los comentarios nunca se borran: son la trazabilidad de cada orden.
"""
import sqlite3
from datetime import datetime, date

import pandas as pd

ESTADOS_SEG = ["Pendiente", "Solicitada al proveedor", "Factura en camino", "Observada / en revisión",
               "Esperando conformidad", "Cerrar sin factura"]
SIT_ACTIVA = "Activa"
SIT_FACTURADA = "Facturada"
SIT_FUERA = "Fuera del reporte"
SIT_SIN_FACT = "Cerrada sin factura"     # decisión del usuario: no se facturará
SIT_CERRADA_SAP = "OC cerrada sin factura"     # la OC se cerró en SAP sin facturar
SIT_CANCELADA_SAP = "OC cancelada sin factura"  # la OC se canceló en SAP sin facturar
# Estado de la OC en SAP que la saca del seguimiento -> situación con la que se cierra
ESTADOS_FINALES = {"Cerrado": SIT_CERRADA_SAP, "Cancelado": SIT_CANCELADA_SAP}
CERRADAS_AUTO = (SIT_FACTURADA, SIT_FUERA, SIT_CERRADA_SAP, SIT_CANCELADA_SAP)

SCHEMA = """
CREATE TABLE IF NOT EXISTS ordenes (
    numero_oc INTEGER PRIMARY KEY,
    serie TEXT, usuario TEXT, proveedor TEXT, ruc TEXT, mon TEXT, total REAL,
    fecha_oc TEXT, estado_oc TEXT,
    situacion TEXT NOT NULL DEFAULT 'Activa',
    estado_seg TEXT NOT NULL DEFAULT 'Pendiente',
    responsable TEXT DEFAULT '', fecha_compromiso TEXT DEFAULT '',
    primer_visto TEXT, ultimo_visto TEXT,
    fecha_cierre TEXT DEFAULT '', facturas TEXT DEFAULT ''
);
CREATE TABLE IF NOT EXISTS comentarios (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    numero_oc INTEGER NOT NULL, fecha TEXT NOT NULL, autor TEXT, texto TEXT NOT NULL,
    tipo TEXT NOT NULL DEFAULT 'usuario'
);
CREATE INDEX IF NOT EXISTS ix_com_oc ON comentarios(numero_oc, fecha);
CREATE TABLE IF NOT EXISTS facturas (
    ruc TEXT NOT NULL, n_factura TEXT NOT NULL,
    numero_oc INTEGER, serie TEXT, usuario TEXT, proveedor TEXT, estado_oc TEXT, fecha_oc TEXT,
    tipo TEXT, estado_factura TEXT, mon TEXT, ft_sin_imp REAL, imp_pendiente REAL,
    emision TEXT, vencimiento TEXT, fecha_pago TEXT,
    situacion TEXT NOT NULL DEFAULT 'Vigente',
    primer_visto TEXT, ultimo_visto TEXT, pagada_detectada TEXT DEFAULT '',
    PRIMARY KEY (ruc, n_factura)
);
CREATE INDEX IF NOT EXISTS ix_fact_oc ON facturas(numero_oc);
CREATE TABLE IF NOT EXISTS cargas (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    fecha_reporte TEXT, archivo TEXT, fecha_carga TEXT, salida TEXT,
    activas INTEGER, nuevas INTEGER, facturadas INTEGER, fuera INTEGER
);
"""


def conectar(path):
    con = sqlite3.connect(path)
    con.row_factory = sqlite3.Row
    con.executescript(SCHEMA)
    # Limpieza única: OCs que entraron ya Cerradas/Canceladas con reglas anteriores y nadie trabajó
    sin_trabajo = """SELECT numero_oc FROM ordenes o WHERE estado_oc IN ('Cerrado','Cancelado') AND situacion='Activa'
                     AND NOT EXISTS (SELECT 1 FROM comentarios c WHERE c.numero_oc=o.numero_oc AND c.tipo<>'sistema')"""
    con.execute(f"DELETE FROM comentarios WHERE numero_oc IN ({sin_trabajo})")
    con.execute(f"DELETE FROM ordenes WHERE numero_oc IN ({sin_trabajo})")
    con.commit()
    return con


def _ahora():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def _iso(d):
    if d is None or (not isinstance(d, str) and pd.isna(d)):
        return ""
    return d.strftime("%Y-%m-%d") if hasattr(d, "strftime") else str(d)


def comentar(con, oc, texto, autor, tipo="usuario", fecha=None):
    con.execute("INSERT INTO comentarios(numero_oc, fecha, autor, texto, tipo) VALUES (?,?,?,?,?)",
                (int(oc), fecha or _ahora(), autor, texto.strip(), tipo))


def ultima_carga(con):
    return con.execute("SELECT * FROM cargas ORDER BY fecha_reporte DESC, id DESC LIMIT 1").fetchone()


def sincronizar(con, df, fecha_reporte, archivo):
    """Actualiza la base con el reporte limpio `df` (salida de procesar_sap.limpiar)."""
    rep = fecha_reporte.strftime("%Y-%m-%d %H:%M:%S")
    ult = ultima_carga(con)
    if ult is not None and rep < ult["fecha_reporte"]:
        return {"omitida": True, "motivo": f"El reporte ({fecha_reporte:%d/%m/%Y %H:%M}) es anterior a la última carga "
                                           f"({ult['fecha_reporte']}). No se actualizó el seguimiento."}
    fecha_ev = f"{rep[:10]} {rep[11:]}"
    rep_txt = f"{fecha_reporte:%d/%m/%Y}"

    fact = df[df["N° Factura"] != ""].groupby("Numero OC")["N° Factura"].apply(lambda s: ", ".join(s))
    cab = df.drop_duplicates("Numero OC").set_index("Numero OC")
    seguibles = [oc for oc in cab.index if oc not in fact.index and cab.loc[oc, "Estado"] not in ESTADOS_FINALES]
    existentes = {r["numero_oc"]: r for r in con.execute("SELECT * FROM ordenes")}
    res = {"omitida": False, "nuevas": 0, "facturadas": 0, "fuera": 0, "cerradas_sap": 0, "canceladas_sap": 0,
           "reabiertas": 0}

    for oc in seguibles:
        c = cab.loc[oc]
        snap = (c["Serie"], c["Usuario"], c["Proveedor"], c["RUC"], c["Mon"], float(c["Total Sin IGV"] or 0),
                _iso(c["Fecha"]), c["Estado"])
        prev = existentes.get(oc)
        if prev is None:
            con.execute("""INSERT INTO ordenes(numero_oc, serie, usuario, proveedor, ruc, mon, total, fecha_oc, estado_oc,
                           situacion, estado_seg, primer_visto, ultimo_visto) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                        (int(oc), *snap, SIT_ACTIVA, "Pendiente", rep, rep))
            comentar(con, oc, f"Ingresa a seguimiento: OC sin factura (Estado OC: {c['Estado']}) en reporte {rep_txt}.",
                     "Sistema", "sistema", fecha_ev)
            res["nuevas"] += 1
            continue
        if prev["estado_oc"] != c["Estado"]:
            comentar(con, oc, f"Estado de la OC cambió de {prev['estado_oc']} a {c['Estado']} (reporte {rep_txt}).",
                     "Sistema", "sistema", fecha_ev)
        situacion, cierre = prev["situacion"], prev["fecha_cierre"]
        if situacion in CERRADAS_AUTO:
            situacion, cierre = SIT_ACTIVA, ""
            comentar(con, oc, f"Vuelve a seguimiento: la OC aparece sin factura vigente en reporte {rep_txt}.",
                     "Sistema", "sistema", fecha_ev)
            res["reabiertas"] += 1
        con.execute("""UPDATE ordenes SET serie=?, usuario=?, proveedor=?, ruc=?, mon=?, total=?, fecha_oc=?, estado_oc=?,
                       situacion=?, fecha_cierre=?, facturas='', ultimo_visto=? WHERE numero_oc=?""",
                    (*snap, situacion, cierre, rep, int(oc)))

    recien_facturadas = set()
    for oc, prev in existentes.items():
        if prev["situacion"] in CERRADAS_AUTO:
            continue
        if oc in fact.index:
            recien_facturadas.add(oc)
            con.execute("UPDATE ordenes SET situacion=?, fecha_cierre=?, facturas=?, ultimo_visto=? WHERE numero_oc=?",
                        (SIT_FACTURADA, rep, fact[oc], rep, oc))
            comentar(con, oc, f"Factura cargada en la OC: {fact[oc]} (detectado en reporte {rep_txt}). Seguimiento cerrado.",
                     "Sistema", "sistema", fecha_ev)
            res["facturadas"] += 1
        elif oc not in cab.index:
            con.execute("UPDATE ordenes SET situacion=?, fecha_cierre=? WHERE numero_oc=?", (SIT_FUERA, rep, oc))
            comentar(con, oc, f"La OC ya no aparece en el reporte {rep_txt}. Seguimiento cerrado.",
                     "Sistema", "sistema", fecha_ev)
            res["fuera"] += 1
        elif cab.loc[oc, "Estado"] in ESTADOS_FINALES:
            est = cab.loc[oc, "Estado"]
            con.execute("UPDATE ordenes SET situacion=?, estado_oc=?, fecha_cierre=?, ultimo_visto=? WHERE numero_oc=?",
                        (ESTADOS_FINALES[est], est, rep, rep, oc))
            verbo = "se cerró" if est == "Cerrado" else "se canceló"
            comentar(con, oc, f"La OC {verbo} en SAP sin factura (reporte {rep_txt}): ya no se facturará. Seguimiento cerrado.",
                     "Sistema", "sistema", fecha_ev)
            res["cerradas_sap" if est == "Cerrado" else "canceladas_sap"] += 1

    _sincronizar_facturas(con, df, rep, rep_txt, fecha_ev, recien_facturadas, res)
    res["activas"] = con.execute("SELECT COUNT(*) FROM ordenes WHERE situacion=?", (SIT_ACTIVA,)).fetchone()[0]
    con.execute("INSERT INTO cargas(fecha_reporte, archivo, fecha_carga, activas, nuevas, facturadas, fuera) VALUES (?,?,?,?,?,?,?)",
                (rep, archivo, _ahora(), res["activas"], res["nuevas"], res["facturadas"], res["fuera"]))
    con.commit()
    return res


FACT_VIGENTE = "Vigente"
FACT_FUERA = "Fuera del reporte"


def _sincronizar_facturas(con, df, rep, rep_txt, fecha_ev, recien_facturadas, res):
    """Guarda una foto de cada factura (RUC + N° Factura) y registra en el historial de la OC
    las facturas nuevas, los cambios de estado (p. ej. Pendiente de pago -> Pagado) y las que desaparecen.
    En la primera carga solo se toma la foto, sin generar comentarios."""
    primera = con.execute("SELECT COUNT(*) FROM facturas").fetchone()[0] == 0
    previas = {(r["ruc"], r["n_factura"]): r for r in con.execute("SELECT * FROM facturas")}
    vistas = set()
    res.update(facturas_nuevas=0, facturas_pagadas=0)
    for _, r in df[df["N° Factura"] != ""].iterrows():
        clave = (r["RUC"], r["N° Factura"])
        vistas.add(clave)
        oc = int(r["Numero OC"])
        snap = (oc, r["Serie"], r["Usuario"], r["Proveedor"], r["Estado"], _iso(r["Fecha"]), r["Tipo"],
                r["Estado Factura"], r["Mon"], None if pd.isna(r["FT Sin Imp."]) else float(r["FT Sin Imp."]),
                None if pd.isna(r["Imp. Pendiente"]) else float(r["Imp. Pendiente"]),
                _iso(r["Emisión FT"]), _iso(r["Venc. FT"]), _iso(r["Fecha Pago"]))
        prev = previas.get(clave)
        if prev is None:
            con.execute("""INSERT INTO facturas(numero_oc, serie, usuario, proveedor, estado_oc, fecha_oc, tipo, estado_factura,
                           mon, ft_sin_imp, imp_pendiente, emision, vencimiento, fecha_pago, ruc, n_factura,
                           situacion, primer_visto, ultimo_visto) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                        (*snap, *clave, FACT_VIGENTE, rep, rep))
            if not primera:
                res["facturas_nuevas"] += 1
                if oc not in recien_facturadas:
                    comentar(con, oc, f"Nueva factura {clave[1]} registrada en la OC ({r['Estado Factura'] or 'sin estado'}) "
                                      f"en reporte {rep_txt}.", "Sistema", "sistema", fecha_ev)
            continue
        pagada = prev["pagada_detectada"]
        if prev["estado_factura"] != r["Estado Factura"]:
            if r["Estado Factura"] == "Pagado":
                fp = f" el {_iso(r['Fecha Pago'])[8:10]}/{_iso(r['Fecha Pago'])[5:7]}/{_iso(r['Fecha Pago'])[:4]}" if _iso(r["Fecha Pago"]) else ""
                comentar(con, oc, f"Factura {clave[1]} pagada{fp} (detectado en reporte {rep_txt}).", "Sistema", "sistema", fecha_ev)
                pagada = rep
                res["facturas_pagadas"] += 1
            else:
                comentar(con, oc, f"Factura {clave[1]}: estado cambió de {prev['estado_factura'] or '(vacío)'} a "
                                  f"{r['Estado Factura'] or '(vacío)'} (reporte {rep_txt}).", "Sistema", "sistema", fecha_ev)
        if prev["situacion"] == FACT_FUERA:
            comentar(con, oc, f"Factura {clave[1]} vuelve a aparecer en el reporte {rep_txt}.", "Sistema", "sistema", fecha_ev)
        con.execute("""UPDATE facturas SET numero_oc=?, serie=?, usuario=?, proveedor=?, estado_oc=?, fecha_oc=?, tipo=?,
                       estado_factura=?, mon=?, ft_sin_imp=?, imp_pendiente=?, emision=?, vencimiento=?, fecha_pago=?,
                       situacion=?, ultimo_visto=?, pagada_detectada=? WHERE ruc=? AND n_factura=?""",
                    (*snap, FACT_VIGENTE, rep, pagada, *clave))
    for clave, prev in previas.items():
        if clave not in vistas and prev["situacion"] == FACT_VIGENTE:
            con.execute("UPDATE facturas SET situacion=? WHERE ruc=? AND n_factura=?", (FACT_FUERA, *clave))
            comentar(con, prev["numero_oc"], f"Factura {clave[1]} ya no aparece en el reporte {rep_txt} (¿anulada?).",
                     "Sistema", "sistema", fecha_ev)


def listar_facturas(con):
    q = """SELECT f.*,
             (SELECT texto FROM comentarios c WHERE c.numero_oc=f.numero_oc AND c.tipo='usuario' ORDER BY fecha DESC, id DESC LIMIT 1) AS ult_com,
             (SELECT fecha FROM comentarios c WHERE c.numero_oc=f.numero_oc AND c.tipo='usuario' ORDER BY fecha DESC, id DESC LIMIT 1) AS ult_fecha,
             (SELECT COUNT(*) FROM comentarios c WHERE c.numero_oc=f.numero_oc AND c.tipo='usuario') AS n_com
           FROM facturas f"""
    return [dict(r) for r in con.execute(q)]


def ultimos_comentarios(con):
    """{numero_oc: (fecha, texto)} con el último comentario de usuario de cada OC."""
    q = """SELECT numero_oc, fecha, texto FROM comentarios c WHERE tipo='usuario' AND id = (
             SELECT id FROM comentarios c2 WHERE c2.numero_oc=c.numero_oc AND c2.tipo='usuario' ORDER BY fecha DESC, id DESC LIMIT 1)"""
    return {r["numero_oc"]: (r["fecha"], r["texto"]) for r in con.execute(q)}


def actualizar_orden(con, oc, estado_seg, responsable, fecha_compromiso, autor):
    """Guarda estado/responsable/compromiso y deja constancia de cada cambio en el historial."""
    prev = con.execute("SELECT * FROM ordenes WHERE numero_oc=?", (oc,)).fetchone()
    cambios = []
    if prev["estado_seg"] != estado_seg:
        cambios.append(f"Estado de seguimiento: {prev['estado_seg']} → {estado_seg}")
    if (prev["responsable"] or "") != responsable:
        cambios.append(f"Responsable: {prev['responsable'] or '(vacío)'} → {responsable or '(vacío)'}")
    if (prev["fecha_compromiso"] or "") != fecha_compromiso:
        f = lambda s: datetime.strptime(s, "%Y-%m-%d").strftime("%d/%m/%Y") if s else "(vacío)"
        cambios.append(f"Fecha compromiso: {f(prev['fecha_compromiso'])} → {f(fecha_compromiso)}")
    if not cambios:
        return False
    situacion = prev["situacion"]
    if prev["situacion"] in (SIT_ACTIVA, SIT_SIN_FACT):
        situacion = SIT_SIN_FACT if estado_seg == "Cerrar sin factura" else SIT_ACTIVA
    con.execute("UPDATE ordenes SET estado_seg=?, responsable=?, fecha_compromiso=?, situacion=?, fecha_cierre=? WHERE numero_oc=?",
                (estado_seg, responsable, fecha_compromiso, situacion,
                 _ahora() if situacion == SIT_SIN_FACT else ("" if situacion == SIT_ACTIVA else prev["fecha_cierre"]), oc))
    comentar(con, oc, "; ".join(cambios), autor, "cambio")
    con.commit()
    return True


def listar(con):
    q = """SELECT o.*,
             (SELECT texto FROM comentarios c WHERE c.numero_oc=o.numero_oc AND c.tipo='usuario' ORDER BY fecha DESC, id DESC LIMIT 1) AS ult_com,
             (SELECT fecha FROM comentarios c WHERE c.numero_oc=o.numero_oc AND c.tipo='usuario' ORDER BY fecha DESC, id DESC LIMIT 1) AS ult_fecha,
             (SELECT COUNT(*) FROM comentarios c WHERE c.numero_oc=o.numero_oc AND c.tipo='usuario') AS n_com
           FROM ordenes o"""
    return [dict(r) for r in con.execute(q)]


def historial(con, oc):
    return [dict(r) for r in con.execute("SELECT * FROM comentarios WHERE numero_oc=? ORDER BY fecha, id", (oc,))]


def tabla_para_excel(con):
    """Filas para la hoja 'Seguimiento' del control (activas primero)."""
    filas = listar(con)
    hist = {}
    for r in con.execute("SELECT * FROM comentarios ORDER BY fecha, id"):
        f = datetime.strptime(r["fecha"], "%Y-%m-%d %H:%M:%S").strftime("%d/%m/%Y %H:%M")
        hist.setdefault(r["numero_oc"], []).append(f"[{f}] {r['autor']}: {r['texto']}")
    orden = {SIT_ACTIVA: 0, SIT_SIN_FACT: 1, SIT_FACTURADA: 2, SIT_CERRADA_SAP: 3, SIT_CANCELADA_SAP: 4, SIT_FUERA: 5}
    filas.sort(key=lambda r: (orden.get(r["situacion"], 9), r["fecha_oc"] or ""))
    for r in filas:
        r["historial"] = "\n".join(hist.get(r["numero_oc"], []))
    return filas


def dias_desde(iso):
    if not iso:
        return None
    return (date.today() - datetime.strptime(iso[:10], "%Y-%m-%d").date()).days
