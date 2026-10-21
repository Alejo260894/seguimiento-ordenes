"""
Seguimiento de requerimientos (solicitudes del Accede / correo) desde la solicitud hasta el pago.

- Base: tablas req_* dentro de seguimiento.db. La app es la fuente; el Excel se genera (exportar_excel).
- Catálogos: proyectos (hoja Consulta1) y proveedores (hoja PROVEEDORES) del Excel
  "SEGUIMIENTO DE REQUERIMIENTOS Y ORDENES", más los proveedores del reporte SAP.
- SAP: con el N° de orden se completan los campos vacíos (emisión/aprobación de OC, proveedor, monto,
  moneda, factura, contabilización y fecha de pago). Lo que se llena desde SAP queda marcado.
- Etapa: se calcula según las fechas registradas (ver ETAPAS).
"""
import json
import os
import shutil
import tempfile
from datetime import date, datetime

import pandas as pd

# Campos de un requerimiento: (columna en la base, encabezado en el Excel CONTROL, tipo)
CAMPOS = [
    ("id", "COD", "int"), ("proyecto", "PROY2", "txt"), ("referencia", "REFERENCIA", "txt"),
    ("tipo", "TIPO", "txt"), ("proveedor", "PROVEEDOR", "txt"), ("f_solic", "F. Solic.", "fecha"),
    ("estatus", "ESTATUS", "txt"), ("fecha_inicio", "FECHA INICIO", "fecha"), ("fecha_fin", "FECHA FIN", "fecha"),
    ("monto", "MONTO", "num"), ("moneda", "MONEDA", "txt"),
    ("envio_valo", "ENVIO DE VALORIZACION", "fecha"), ("envio_aprob_valo", "ENVIO PARA APROBACION VALO", "fecha"),
    ("aprob_valo", "FECHA DE APROBACION VALO", "fecha"),
    ("emision_oc", "EMISION DE ORDEN", "fecha"), ("aprob_oc", "APROBACION DE ORDEN", "fecha"),
    ("envio_oc", "ENVIO DE ORDEN", "fecha"), ("orden", "ORDEN", "int"),
    ("recepcion_fact", "RECEPCION DE FACTURA", "fecha"), ("envio_aprob_fact", "ENVIO DE APROBACION DE FACTURA", "fecha"),
    ("validacion_fact", "VALIDACION DE FACTURA", "fecha"), ("num_factura", "NUMERO DE FACTURA", "txt"),
    ("envio_pago", "ENVIO PARA PAGO", "fecha"), ("fecha_pago", "FECHA DE PAGO", "fecha"),
]
CAMPOS_FECHA = [c for c, _, t in CAMPOS if t == "fecha"]
ETIQUETAS = {
    "proyecto": "Proyecto", "referencia": "Referencia", "tipo": "Tipo", "proveedor": "Proveedor",
    "f_solic": "Fecha de solicitud", "estatus": "Estatus", "fecha_inicio": "Fecha inicio", "fecha_fin": "Fecha fin",
    "monto": "Monto", "moneda": "Moneda", "envio_valo": "Envío de valorización",
    "envio_aprob_valo": "Envío para aprobación de valo", "aprob_valo": "Aprobación de valo",
    "emision_oc": "Emisión de orden", "aprob_oc": "Aprobación de orden", "envio_oc": "Envío de orden",
    "orden": "N° de orden", "recepcion_fact": "Recepción de factura", "envio_aprob_fact": "Envío de aprobación de factura",
    "validacion_fact": "Validación de factura", "num_factura": "N° de factura", "envio_pago": "Envío para pago",
    "fecha_pago": "Fecha de pago", "notas": "Notas",
}
TIPOS = ["RQ", "VALO", "VUELO", "RH"]
MONEDAS = ["PEN", "USD", "CAD", "AUD", "GBP", "EUR", "ZAR"]
MON_SAP = {"SOL": "PEN"}
ESTATUS_CIERRE = {"COMPLETO": "Completo", "ANULADO": "Anulado"}

# Etapas en orden; cada una empieza con la fecha indicada (la más avanzada registrada manda)
ETAPAS = [
    ("Por valorizar", "f_solic"),
    ("Valo por enviar a aprobación", "envio_valo"),
    ("Valo en aprobación", "envio_aprob_valo"),
    ("OC por emitir", "aprob_valo"),
    ("OC por aprobar", "emision_oc"),
    ("OC por enviar", "aprob_oc"),
    ("Pendiente de factura", "envio_oc"),
    ("Factura recibida", "recepcion_fact"),
    ("Factura en validación", "envio_aprob_fact"),
    ("Por enviar a pago", "validacion_fact"),
    ("Pendiente de pago", "envio_pago"),
    ("Pagado", "fecha_pago"),
]
NOMBRES_ETAPA = [e for e, _ in ETAPAS]
ETAPAS_HASTA_OC = NOMBRES_ETAPA[:6]          # el seguimiento principal: hasta que se envía la orden
# Cierres: manuales (estatus COMPLETO/ANULADO) o porque la OC se cerró/canceló en SAP sin factura
SAP_CIERRE = {"Cerrado": "OC cerrada sin factura", "Cancelado": "OC cancelada"}
CERRADAS = ["Pagado", "Completo", "Anulado", *SAP_CIERRE.values()]

TRAMOS = [
    ("Solicitud → Aprobación de valo", "f_solic", "aprob_valo"),
    ("Aprobación de valo → Emisión de OC", "aprob_valo", "emision_oc"),
    ("Emisión de OC → Aprobación de OC", "emision_oc", "aprob_oc"),
    ("Aprobación de OC → Envío de OC", "aprob_oc", "envio_oc"),
    ("Solicitud → Envío de OC (total)", "f_solic", "envio_oc"),
    ("Envío de OC → Recepción de factura", "envio_oc", "recepcion_fact"),
    ("Recepción de factura → Envío para pago", "recepcion_fact", "envio_pago"),
    ("Envío para pago → Pago", "envio_pago", "fecha_pago"),
    ("Solicitud → Pago (total)", "f_solic", "fecha_pago"),
]

SCHEMA = """
CREATE TABLE IF NOT EXISTS req (
    id INTEGER PRIMARY KEY,
    proyecto TEXT DEFAULT '', referencia TEXT DEFAULT '', tipo TEXT DEFAULT '', proveedor TEXT DEFAULT '',
    f_solic TEXT DEFAULT '', estatus TEXT DEFAULT '', fecha_inicio TEXT DEFAULT '', fecha_fin TEXT DEFAULT '',
    monto REAL, moneda TEXT DEFAULT '',
    envio_valo TEXT DEFAULT '', envio_aprob_valo TEXT DEFAULT '', aprob_valo TEXT DEFAULT '',
    emision_oc TEXT DEFAULT '', aprob_oc TEXT DEFAULT '', envio_oc TEXT DEFAULT '', orden INTEGER,
    recepcion_fact TEXT DEFAULT '', envio_aprob_fact TEXT DEFAULT '', validacion_fact TEXT DEFAULT '',
    num_factura TEXT DEFAULT '', envio_pago TEXT DEFAULT '', fecha_pago TEXT DEFAULT '',
    notas TEXT DEFAULT '', desde_sap TEXT DEFAULT '[]', creado TEXT, modificado TEXT
);
CREATE TABLE IF NOT EXISTS req_eventos (
    id INTEGER PRIMARY KEY AUTOINCREMENT, req_id INTEGER NOT NULL, fecha TEXT NOT NULL,
    autor TEXT, texto TEXT NOT NULL, tipo TEXT NOT NULL DEFAULT 'usuario'
);
CREATE INDEX IF NOT EXISTS ix_req_ev ON req_eventos(req_id, fecha);
CREATE TABLE IF NOT EXISTS cat_proyectos (
    codigo TEXT PRIMARY KEY, nombre TEXT, nombre_corto TEXT, estado TEXT, cliente TEXT,
    responsable TEXT, controller1 TEXT, controller2 TEXT
);
CREATE TABLE IF NOT EXISTS cat_proveedores (
    nombre TEXT PRIMARY KEY, ruc TEXT, origen TEXT
);
CREATE TABLE IF NOT EXISTS sap_oc (
    numero_oc INTEGER PRIMARY KEY, fecha TEXT, aprobacion TEXT, usuario TEXT, proveedor TEXT, ruc TEXT,
    mon TEXT, total REAL, estado TEXT
);
CREATE TABLE IF NOT EXISTS sap_fact (
    numero_oc INTEGER, n_factura TEXT, estado TEXT, contab TEXT, emision TEXT, fecha_pago TEXT
);
CREATE INDEX IF NOT EXISTS ix_sap_fact ON sap_fact(numero_oc);
"""


def preparar(con):
    con.executescript(SCHEMA)
    con.commit()


def _ahora():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def _iso(v):
    """Fecha (datetime/Timestamp/str) -> 'YYYY-MM-DD' o ''."""
    if v is None or (not isinstance(v, str) and pd.isna(v)):
        return ""
    if isinstance(v, str):
        v = v.strip()
        for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%Y-%m-%d %H:%M:%S"):
            try:
                return datetime.strptime(v, fmt).strftime("%Y-%m-%d")
            except ValueError:
                pass
        return ""
    return v.strftime("%Y-%m-%d")


def _txt(v):
    if v is None or (not isinstance(v, str) and pd.isna(v)):
        return ""
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    return " ".join(str(v).split())


def _num(v):
    try:
        f = float(v)
        return None if pd.isna(f) else f
    except (TypeError, ValueError):
        return None


def _int(v):
    n = _num(v)
    return None if n is None else int(n)


def fmt(iso):
    return f"{iso[8:10]}/{iso[5:7]}/{iso[:4]}" if iso else ""


# ---------------------------------------------------------------------------
# Lectura del Excel de seguimiento (puede estar abierto en Excel: se copia antes de leer)
# ---------------------------------------------------------------------------
def _copia_lectura(ruta):
    tmp = os.path.join(tempfile.gettempdir(), f"req_lectura_{os.getpid()}.xlsx")
    try:
        shutil.copy2(ruta, tmp)
    except PermissionError:
        with open(ruta, "rb") as src, open(tmp, "wb") as dst:
            shutil.copyfileobj(src, dst)
    return tmp


def actualizar_catalogos(con, ruta_excel, sap_df=None):
    """Recarga proyectos (Consulta1) y proveedores (PROVEEDORES + SAP)."""
    tmp = _copia_lectura(ruta_excel)
    try:
        q = pd.read_excel(tmp, sheet_name="Consulta1", dtype=object)
        pv = pd.read_excel(tmp, sheet_name="PROVEEDORES", dtype=object)
    finally:
        os.remove(tmp)
    q.columns = [_txt(c).upper() for c in q.columns]
    col = lambda *nombres: next(c for c in q.columns if any(c.startswith(n) for n in nombres))
    filas = []
    for _, r in q.iterrows():
        cod = _txt(r[col("CÓDIGO", "CODIGO", "C")])
        if cod:
            filas.append((cod, _txt(r[col("NOMBRE")]), _txt(r[col("NOMBRE CORTO")]), _txt(r[col("ESTADO")]),
                          _txt(r[col("CLIENTE")]), _txt(r[col("RESPONSABLE")]), _txt(r[col("CONTROL PLANNER 1")]),
                          _txt(r[col("CONTROL PLANNER 2")])))
    con.execute("DELETE FROM cat_proyectos")
    con.executemany("INSERT OR REPLACE INTO cat_proyectos VALUES (?,?,?,?,?,?,?,?)", filas)
    prov = {}
    for _, r in pv.iterrows():
        n = _txt(r.get("RazonSocial"))
        if n:
            prov[n.upper()] = (n, _txt(r.get("RUC")), "Excel")
    if sap_df is not None:
        for (n, ruc) in sap_df[["Proveedor", "RUC"]].drop_duplicates().itertuples(index=False):
            prov.setdefault(n.upper(), (n, ruc, "SAP"))
    else:
        for n, ruc in con.execute("SELECT DISTINCT proveedor, ruc FROM sap_oc"):
            if n:
                prov.setdefault(n.upper(), (n, ruc, "SAP"))
    con.execute("DELETE FROM cat_proveedores")
    con.executemany("INSERT OR REPLACE INTO cat_proveedores VALUES (?,?,?)", list(prov.values()))
    con.commit()
    return len(filas), len(prov)


def importar_control(con, ruta_excel, autor="Importación"):
    """Importa la tabla CONTROL del Excel. Solo agrega los COD que no existan en la base."""
    tmp = _copia_lectura(ruta_excel)
    try:
        c = pd.read_excel(tmp, sheet_name="CONTROL", header=1, dtype=object)
    finally:
        os.remove(tmp)
    c.columns = [_txt(x) for x in c.columns]
    existentes = {r[0] for r in con.execute("SELECT id FROM req")}
    nuevos = 0
    for _, r in c.iterrows():
        cod = _int(r.get("COD"))
        if cod is None or cod in existentes or not any(_txt(r.get(h)) for _, h, _ in CAMPOS[1:]):
            continue
        valores = {}
        for campo, enc, tipo in CAMPOS:
            v = r.get(enc)
            valores[campo] = (_iso(v) if tipo == "fecha" else _num(v) if tipo == "num"
                              else _int(v) if tipo == "int" else _txt(v))
        valores["id"] = cod
        valores["estatus"] = valores["estatus"].upper()
        cols = list(valores) + ["creado", "modificado"]
        con.execute(f"INSERT INTO req({','.join(cols)}) VALUES ({','.join('?' * len(cols))})",
                    [valores[k] for k in valores] + [_ahora(), _ahora()])
        evento(con, cod, "Importado desde el Excel de seguimiento.", autor, "sistema")
        nuevos += 1
    con.commit()
    return nuevos


# ---------------------------------------------------------------------------
# SAP
# ---------------------------------------------------------------------------
def guardar_snapshot_sap(con, df):
    """Foto del último reporte SAP limpio (procesar_sap.limpiar), para completar requerimientos."""
    cab = df.drop_duplicates("Numero OC")
    con.execute("DELETE FROM sap_oc")
    con.executemany("INSERT INTO sap_oc VALUES (?,?,?,?,?,?,?,?,?)", [
        (int(r["Numero OC"]), _iso(r["Fecha"]), _iso(r["Aprobación"]), r["Usuario"], r["Proveedor"], r["RUC"],
         MON_SAP.get(r["Mon"], r["Mon"]), _num(r["Total Sin IGV"]), r["Estado"]) for _, r in cab.iterrows()])
    con.execute("DELETE FROM sap_fact")
    fx = df[df["N° Factura"] != ""]
    con.executemany("INSERT INTO sap_fact VALUES (?,?,?,?,?,?)", [
        (int(r["Numero OC"]), r["N° Factura"], r["Estado Factura"], _iso(r["Contab. FT"]), _iso(r["Emisión FT"]),
         _iso(r["Fecha Pago"])) for _, r in fx.iterrows()])
    con.commit()


def datos_sap(con, orden):
    """Valores que SAP aporta para una orden (o None si no está en el último reporte)."""
    if not orden:
        return None
    oc = con.execute("SELECT * FROM sap_oc WHERE numero_oc=?", (int(orden),)).fetchone()
    if oc is None:
        return None
    fac = [dict(r) for r in con.execute("SELECT * FROM sap_fact WHERE numero_oc=? ORDER BY contab", (int(orden),))]
    pagadas = [f for f in fac if f["estado"] == "Pagado" and f["fecha_pago"]]
    d = {"emision_oc": oc["fecha"], "aprob_oc": oc["aprobacion"], "proveedor": oc["proveedor"],
         "monto": oc["total"], "moneda": oc["mon"], "estado_oc": oc["estado"], "facturas": fac}
    if fac:
        d["num_factura"] = ", ".join(f["n_factura"] for f in fac)
        d["recepcion_fact"] = min(f["contab"] for f in fac if f["contab"]) if any(f["contab"] for f in fac) else ""
        # Fecha de pago: solo cuando TODAS las facturas de la OC están pagadas
        if len(pagadas) == len(fac):
            d["fecha_pago"] = max(f["fecha_pago"] for f in pagadas)
    return d


# Campos que se completan desde SAP cuando están vacíos
CAMPOS_SAP = ["emision_oc", "aprob_oc", "proveedor", "monto", "moneda", "num_factura", "recepcion_fact", "fecha_pago"]
# Campos vinculados: si la orden está en SAP, SAP manda siempre (lo aprobado en SAP es lo correcto)
CAMPOS_VINCULADOS = ["monto", "moneda"]


def _distinto(campo, a, b):
    if campo == "monto":
        return a is None or b is None or abs(float(a) - float(b)) > 0.005
    return (a or "") != (b or "")


def completar_desde_sap(con, req_id=None, autor="Sistema"):
    """Actualiza monto y moneda con SAP y rellena los demás campos vacíos.
    Devuelve cuántos requerimientos cambiaron."""
    q = "SELECT * FROM req WHERE orden IS NOT NULL" + (" AND id=?" if req_id else "")
    cambiados = 0
    for r in con.execute(q, (req_id,) if req_id else ()).fetchall():
        d = datos_sap(con, r["orden"])
        if not d:
            continue
        desde = set(json.loads(r["desde_sap"] or "[]"))
        llenados, actualizados = {}, {}
        for campo in CAMPOS_SAP:
            v = d.get(campo)
            if v in (None, ""):
                continue
            if campo in CAMPOS_VINCULADOS:
                desde.add(campo)
                if _distinto(campo, r[campo], v):
                    (actualizados if r[campo] not in (None, "") else llenados)[campo] = v
            elif r[campo] in (None, ""):
                llenados[campo] = v
                desde.add(campo)
        cambios = {**llenados, **actualizados}
        if cambios or sorted(desde) != json.loads(r["desde_sap"] or "[]"):
            sets = "".join(f"{k}=?, " for k in cambios)
            con.execute(f"UPDATE req SET {sets}desde_sap=?, modificado=? WHERE id=?",
                        [*cambios.values(), json.dumps(sorted(desde)), _ahora(), r["id"]])
        ver = lambda k, v: fmt(v) if k in CAMPOS_FECHA else (f"{v:,.2f}" if k == "monto" else v)
        if llenados:
            det = "; ".join(f"{ETIQUETAS[k]}: {ver(k, v)}" for k, v in llenados.items())
            evento(con, r["id"], f"Completado desde SAP (OC {r['orden']}): {det}", autor, "sistema")
        if actualizados:
            det = "; ".join(f"{ETIQUETAS[k]}: {ver(k, r[k])} → {ver(k, v)}" for k, v in actualizados.items())
            evento(con, r["id"], f"Actualizado según lo aprobado en SAP (OC {r['orden']}): {det}", autor, "sistema")
        cambiados += bool(cambios)
    con.commit()
    return cambiados


# ---------------------------------------------------------------------------
# CRUD, etapa e indicadores
# ---------------------------------------------------------------------------
def evento(con, req_id, texto, autor, tipo="usuario"):
    con.execute("INSERT INTO req_eventos(req_id, fecha, autor, texto, tipo) VALUES (?,?,?,?,?)",
                (int(req_id), _ahora(), autor, texto, tipo))


def eventos(con, req_id):
    return [dict(r) for r in con.execute("SELECT * FROM req_eventos WHERE req_id=? ORDER BY fecha, id", (req_id,))]


def siguiente_id(con):
    return (con.execute("SELECT MAX(id) FROM req").fetchone()[0] or 0) + 1


def guardar(con, valores, autor):
    """Crea o actualiza un requerimiento; registra en el historial cada campo que cambia."""
    rid = valores.get("id")
    prev = con.execute("SELECT * FROM req WHERE id=?", (rid,)).fetchone() if rid else None
    campos = [c for c, _, _ in CAMPOS if c != "id"] + ["notas"]
    if prev is None:
        rid = rid or siguiente_id(con)
        cols = ["id"] + campos + ["creado", "modificado"]
        con.execute(f"INSERT INTO req({','.join(cols)}) VALUES ({','.join('?' * len(cols))})",
                    [rid] + [valores.get(c) for c in campos] + [_ahora(), _ahora()])
        evento(con, rid, "Requerimiento registrado.", autor, "cambio")
    else:
        desde = set(json.loads(prev["desde_sap"] or "[]"))
        cambios = []
        for c in campos:
            nuevo, viejo = valores.get(c), prev[c]
            if (nuevo or None) != (viejo or None):
                ver = lambda v: (fmt(v) if c in CAMPOS_FECHA else (f"{v:,.2f}" if c == "monto" and v is not None else v)) or "(vacío)"
                if c != "notas":
                    cambios.append(f"{ETIQUETAS[c]}: {ver(viejo)} → {ver(nuevo)}")
                desde.discard(c)      # editado a mano: ya no es dato de SAP
        if not cambios and (valores.get("notas") or "") == (prev["notas"] or ""):
            return rid
        sets = ", ".join(f"{c}=?" for c in campos)
        con.execute(f"UPDATE req SET {sets}, desde_sap=?, modificado=? WHERE id=?",
                    [valores.get(c) for c in campos] + [json.dumps(sorted(desde)), _ahora(), rid])
        if cambios:
            evento(con, rid, "; ".join(cambios), autor, "cambio")
    con.commit()
    if valores.get("orden"):
        completar_desde_sap(con, rid)
    return rid


def etapa(r):
    """(nombre de etapa, fecha ISO en que empezó). Usa r['sap_estado'] y r['sap_n_fact'] si existen."""
    cierre = ESTATUS_CIERRE.get((r["estatus"] or "").upper())
    if r["fecha_pago"]:
        return "Pagado", r["fecha_pago"]
    if cierre:
        return cierre, r["modificado"][:10] if r.get("modificado") else ""
    if r.get("sap_estado") in SAP_CIERRE and not r.get("sap_n_fact") and not r["num_factura"]:
        return SAP_CIERRE[r["sap_estado"]], ""
    for nombre, campo in reversed(ETAPAS):
        if campo == "orden":
            continue
        if r[campo]:
            return nombre, r[campo]
    if r["orden"]:
        return "OC por aprobar", r["emision_oc"] or r["f_solic"]
    return "Por valorizar", r["f_solic"]


def listar(con):
    proys = {r["codigo"]: dict(r) for r in con.execute("SELECT * FROM cat_proyectos")}
    sap = {r["numero_oc"]: r["estado"] for r in con.execute("SELECT numero_oc, estado FROM sap_oc")}
    n_fact = dict(con.execute("SELECT numero_oc, COUNT(*) FROM sap_fact GROUP BY numero_oc").fetchall())
    # Último comentario de usuario de cada OC (pestañas de seguimiento de OC y facturas)
    com = {}
    if con.execute("SELECT 1 FROM sqlite_master WHERE name='comentarios'").fetchone():
        com = {r[0]: r[1] for r in con.execute(
            """SELECT numero_oc, texto FROM comentarios c WHERE tipo='usuario' AND id=(SELECT id FROM comentarios c2
               WHERE c2.numero_oc=c.numero_oc AND c2.tipo='usuario' ORDER BY fecha DESC, id DESC LIMIT 1)""")}
    filas = []
    hoy = date.today()
    for r in con.execute("SELECT * FROM req ORDER BY id"):
        d = dict(r)
        d["sap_estado"] = sap.get(d["orden"], "") if d["orden"] else ""
        d["sap_n_fact"] = n_fact.get(d["orden"], 0) if d["orden"] else 0
        d["com_oc"] = com.get(d["orden"], "") if d["orden"] else ""
        d["etapa"], ini = etapa(d)
        d["dias_etapa"] = (hoy - datetime.strptime(ini, "%Y-%m-%d").date()).days if ini and d["etapa"] not in CERRADAS else None
        p = proys.get(d["proyecto"], {})
        d["proyecto_nombre"] = p.get("nombre_corto") or p.get("nombre") or ""
        d["controller"] = p.get("controller1", "")
        d["controller2"] = p.get("controller2", "")
        d["jefe"] = p.get("responsable", "")
        d["desde_sap"] = json.loads(d["desde_sap"] or "[]")
        filas.append(d)
    return filas


def dias_entre(a, b):
    if not a or not b:
        return None
    return (datetime.strptime(b, "%Y-%m-%d") - datetime.strptime(a, "%Y-%m-%d")).days


def indicadores(filas):
    """Pendientes por etapa y tiempos del flujo (días) para las filas dadas."""
    pend = []
    for nombre in NOMBRES_ETAPA[:-1]:
        g = [f for f in filas if f["etapa"] == nombre]
        dias = [f["dias_etapa"] for f in g if f["dias_etapa"] is not None]
        montos = {}
        for f in g:
            if f["monto"] and f["moneda"]:
                montos[f["moneda"]] = montos.get(f["moneda"], 0) + f["monto"]
        pend.append(dict(etapa=nombre, n=len(g), prom=(sum(dias) / len(dias)) if dias else None,
                         max=max(dias) if dias else None, montos=montos, hasta_oc=nombre in ETAPAS_HASTA_OC))
    tiempos = []
    for nombre, a, b in TRAMOS:
        v = sorted(d for d in (dias_entre(f[a], f[b]) for f in filas) if d is not None and d >= 0)
        neg = sum(1 for f in filas if (dias_entre(f[a], f[b]) or 0) < 0)
        med = (v[len(v) // 2] if len(v) % 2 else (v[len(v) // 2 - 1] + v[len(v) // 2]) / 2) if v else None
        tiempos.append(dict(tramo=nombre, n=len(v), prom=(sum(v) / len(v)) if v else None, mediana=med,
                            max=v[-1] if v else None, negativos=neg))
    return pend, tiempos


# ---------------------------------------------------------------------------
# Exportación a Excel
# ---------------------------------------------------------------------------
def exportar_excel(con, salida):
    import xlsxwriter
    filas = listar(con)
    wb = xlsxwriter.Workbook(salida, {"default_format_properties": {"font_name": "Arial", "font_size": 10}})
    F = lambda **k: wb.add_format({"font_name": "Arial", "font_size": 10, **k})
    AZUL = "#1F3864"
    hdr = F(bold=True, font_color="white", bg_color=AZUL, text_wrap=True, valign="vcenter")
    grp = F(bold=True, font_color="white", bg_color="#2F5597", align="center")
    fd, fm, fsap = F(num_format="dd/mm/yyyy"), F(num_format="#,##0.00"), F(num_format="dd/mm/yyyy", font_color="#2F5597", italic=True)
    titulo, sub = F(bold=True, font_size=14), F(font_color="#595959")

    # --- CONTROL (mismas columnas que el Excel original + etapa y tiempos)
    ws = wb.add_worksheet("CONTROL")
    extra = ["CONTROLLER", "CONTROLLER2", "JEFE DE PROYECTO", "ETAPA", "DÍAS EN ETAPA", "ESTADO OC SAP",
             "ÚLTIMO COMENTARIO OC", "DATOS DESDE SAP", "NOTAS"]
    tramos = [t[0] for t in TRAMOS]
    enc = [h for _, h, _ in CAMPOS] + extra + [f"DÍAS: {t}" for t in tramos]
    pos = {h: i for i, h in enumerate(enc)}
    grupos = [("DETALLE DE REQUERIMIENTO", "FECHA INICIO", "MONEDA"), ("FLUJO VALORIZACION", "ENVIO DE VALORIZACION", "FECHA DE APROBACION VALO"),
              ("FLUJO ORDEN", "EMISION DE ORDEN", "ORDEN"), ("FLUJO PAGO", "RECEPCION DE FACTURA", "FECHA DE PAGO"),
              ("SEGUIMIENTO", "CONTROLLER", "NOTAS"), ("TIEMPOS (días)", enc[len(CAMPOS) + len(extra)], enc[-1])]
    for g, a, b in grupos:
        ws.merge_range(0, pos[a], 0, pos[b], g, grp)
    n = len(filas)
    datos = []
    for f in filas:
        fila = []
        for campo, _, tipo in CAMPOS:
            v = f[campo]
            fila.append(datetime.strptime(v, "%Y-%m-%d") if tipo == "fecha" and v else (v if v not in ("", None) else None))
        fila += [f["controller"], f["controller2"], f["jefe"], f["etapa"], f["dias_etapa"], f["sap_estado"] or None,
                 f["com_oc"] or None, ", ".join(ETIQUETAS[c] for c in f["desde_sap"]), f["notas"] or None]
        fila += [None] * len(tramos)
        datos.append(fila)
    encabezado = {k: e for k, e, _ in CAMPOS}
    cols = []
    for h in enc:
        c = {"header": h, "header_format": hdr}
        if h.startswith("DÍAS: "):
            _, a, b = next(t for t in TRAMOS if t[0] == h[6:])
            ea, eb = f"[@[{encabezado[a]}]]", f"[@[{encabezado[b]}]]"
            c["formula"] = f'=IF(OR({ea}="",{eb}=""),"",{eb}-{ea})'
        cols.append(c)
    ws.add_table(1, 0, 1 + max(n, 1), len(enc) - 1, {"name": "Tabla1", "columns": cols, "data": datos,
                                                       "style": "Table Style Light 9"})
    for i, (campo, h, tipo) in enumerate(CAMPOS):
        ws.set_column(i, i, 12 if tipo == "fecha" else 10, fd if tipo == "fecha" else fm if tipo == "num" else None)
    ws.set_column(pos["PROVEEDOR"], pos["PROVEEDOR"], 34)
    ws.set_column(pos["REFERENCIA"], pos["REFERENCIA"], 14)
    for h in ("CONTROLLER", "CONTROLLER2", "JEFE DE PROYECTO"):
        ws.set_column(pos[h], pos[h], 26)
    ws.set_column(pos["ETAPA"], pos["ETAPA"], 24)
    ws.set_column(pos["DATOS DESDE SAP"], pos["DATOS DESDE SAP"], 30)
    ws.set_column(pos["ESTADO OC SAP"], pos["ESTADO OC SAP"], 13)
    ws.set_column(pos["ÚLTIMO COMENTARIO OC"], pos["ÚLTIMO COMENTARIO OC"], 40)
    ws.set_column(pos["NOTAS"], pos["NOTAS"], 40)
    ws.set_column(pos["NOTAS"] + 1, len(enc) - 1, 14)
    ws.set_row(1, 45)
    ws.freeze_panes(2, 3)
    # Fechas que vinieron de SAP en azul cursiva
    for r, f in enumerate(filas, start=2):
        for campo in f["desde_sap"]:
            if campo in CAMPOS_FECHA and f[campo]:
                ws.write_datetime(r, next(i for i, (k, _, _) in enumerate(CAMPOS) if k == campo),
                                  datetime.strptime(f[campo], "%Y-%m-%d"), fsap)

    # --- PENDIENTES
    pe = wb.add_worksheet("PENDIENTES")
    pe.hide_gridlines(2)
    pe.write("A1", "Pendientes por etapa", titulo)
    pe.write("A2", "Fórmulas sobre la tabla CONTROL (columna ETAPA y DÍAS EN ETAPA). Las primeras 6 etapas son el "
                   "seguimiento hasta el envío de la orden.", sub)
    pe.write_row(3, 0, ["Etapa", "N°", "Días promedio en la etapa", "Máx. días", "Monto PEN", "Monto USD"], hdr)
    for i, nombre in enumerate(NOMBRES_ETAPA[:-1]):
        r = 4 + i
        pe.write(r, 0, nombre, F(bold=nombre in ETAPAS_HASTA_OC))
        pe.write_formula(r, 1, f'=COUNTIFS(Tabla1[ETAPA],$A{r + 1})')
        pe.write_formula(r, 2, f'=IFERROR(AVERAGEIFS(Tabla1[DÍAS EN ETAPA],Tabla1[ETAPA],$A{r + 1}),"-")', F(num_format="0.0"))
        pe.write_formula(r, 3, f'=IF(B{r + 1}=0,"-",_xlfn.MAXIFS(Tabla1[DÍAS EN ETAPA],Tabla1[ETAPA],$A{r + 1}))')
        pe.write_formula(r, 4, f'=SUMIFS(Tabla1[MONTO],Tabla1[ETAPA],$A{r + 1},Tabla1[MONEDA],"PEN")', F(num_format='#,##0.00;-#,##0.00;"-"'))
        pe.write_formula(r, 5, f'=SUMIFS(Tabla1[MONTO],Tabla1[ETAPA],$A{r + 1},Tabla1[MONEDA],"USD")', F(num_format='#,##0.00;-#,##0.00;"-"'))
    fin = 4 + len(NOMBRES_ETAPA) - 1
    pe.write(fin, 0, "Total pendientes hasta envío de OC", F(bold=True, bg_color="#F2F2F2"))
    pe.write_formula(fin, 1, f"=SUM(B5:B{4 + len(ETAPAS_HASTA_OC)})", F(bold=True, bg_color="#F2F2F2"))
    pe.set_column("A:A", 34)
    pe.set_column("B:F", 16)

    # --- TIEMPOS
    ti = wb.add_worksheet("TIEMPOS")
    ti.hide_gridlines(2)
    ti.write("A1", "Tiempos del flujo (días)", titulo)
    ti.write("A2", "Calculados sobre las columnas DÍAS de la tabla CONTROL; solo cuentan los requerimientos con ambas "
                   "fechas. Los días negativos corresponden a anticipos o regularizaciones (la fecha es anterior a la "
                   "orden): no entran en el promedio y se cuentan aparte.", sub)
    ti.write_row(3, 0, ["Tramo", "N°", "Promedio", "Mediana", "Máximo"] + [f"Promedio {t}" for t in TIPOS]
                 + ["Anticipos / regularizaciones"], hdr)
    for i, t in enumerate(tramos):
        r = 4 + i
        colref = f"Tabla1[DÍAS: {t}]"
        ti.write(r, 0, t, F(bold="total" in t))
        ti.write_formula(r, 1, f'=COUNTIFS({colref},">=0")')
        ti.write_formula(r, 2, f'=IFERROR(AVERAGEIFS({colref},{colref},">=0"),"-")', F(num_format="0.0"))
        ti.write_dynamic_array_formula(r, 3, r, 3, f'=IFERROR(MEDIAN(IF(ISNUMBER({colref}),IF({colref}>=0,{colref}))),"-")',
                                       F(num_format="0.0"))
        ti.write_formula(r, 4, f'=IF(B{r + 1}=0,"-",MAX({colref}))')
        for k, tp in enumerate(TIPOS):
            ti.write_formula(r, 5 + k, f'=IFERROR(AVERAGEIFS({colref},Tabla1[TIPO],"{tp}",{colref},">=0"),"-")',
                             F(num_format="0.0"))
        ti.write_formula(r, 5 + len(TIPOS), f'=COUNTIFS({colref},"<0")', F(font_color="#7F7F7F"))
    ti.set_column("A:A", 42)
    ti.set_column("B:J", 13)
    wb.close()
    return salida
