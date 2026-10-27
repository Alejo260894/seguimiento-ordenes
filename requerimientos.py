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
import numbers
import os
import shutil
import sqlite3
import tempfile
from datetime import date, datetime, timedelta

import pandas as pd

import detalle_sap
import seguimiento as sgm

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
    "fecha_pago": "Fecha de pago", "notas": "Notas", "comprador": "Comprador",
    "monto_ref": "Monto referencial", "tipo_ref": "Tipo de referencia", "f_conformidad": "Fecha de conformidad",
}
# Campos para los KPI de abastecimiento (no existen en el Excel CONTROL original):
# - monto_ref / tipo_ref: precio de referencia (cotización inicial, presupuesto…) -> Ahorro = referencial − monto OC
# - f_conformidad: entrega real o conformidad del servicio -> a tiempo si es ≤ fecha fin (OTIF)
CAMPOS_KPI = ["monto_ref", "tipo_ref", "f_conformidad"]
TIPOS_REF = ["Cotización inicial", "Presupuesto", "Precio histórico", "Precio de mercado", "Licitación"]
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
CREATE TABLE IF NOT EXISTS req_config (clave TEXT PRIMARY KEY, valor TEXT);
"""


def _agregar_columna(con, tabla, columna, tipo):
    if columna not in {r[1] for r in con.execute(f"PRAGMA table_info({tabla})")}:
        con.execute(f"ALTER TABLE {tabla} ADD COLUMN {columna} {tipo}")


def preparar(con):
    con.executescript(SCHEMA)
    _agregar_columna(con, "sap_fact", "monto", "REAL")               # FT Sin Imp. de cada factura
    _agregar_columna(con, "sap_fact", "pendiente", "REAL")           # Imp. Pendiente (saldo por pagar)
    _agregar_columna(con, "req", "manual", "TEXT DEFAULT '[]'")      # campos que el usuario vació a propósito
    _agregar_columna(con, "sap_oc", "serie", "TEXT DEFAULT ''")      # OCB/OCS: excepción de nota de entrada
    for col, tipo in (("tipo_oc", "TEXT"), ("estado_logistico", "TEXT"), ("ingresado", "INTEGER"), ("fecha_entrada", "TEXT"),
                      ("pct_atendido", "REAL")):
        _agregar_columna(con, "sap_oc", col, tipo)                   # del reporte de detalle (detalle_sap.enlazar)
    _agregar_columna(con, "req", "monto_ref", "REAL")
    _agregar_columna(con, "req", "tipo_ref", "TEXT DEFAULT ''")
    _agregar_columna(con, "req", "f_conformidad", "TEXT DEFAULT ''")
    if "comprador" not in {r[1] for r in con.execute("PRAGMA table_info(req)")}:
        # Base compartida entre compradores: cada requerimiento tiene su comprador (mismo nombre que «Usuario» en SAP)
        con.execute("ALTER TABLE req ADD COLUMN comprador TEXT DEFAULT ''")
        con.execute("ALTER TABLE req ADD COLUMN cod_excel INTEGER")     # COD original en el Excel importado
        con.execute("""UPDATE req SET comprador=COALESCE((SELECT usuario FROM sap_oc s WHERE s.numero_oc=req.orden), '')""")
        mayor = con.execute("SELECT comprador FROM req WHERE comprador<>'' GROUP BY comprador "
                            "ORDER BY COUNT(*) DESC LIMIT 1").fetchone()
        if mayor:       # hasta ahora la base era de un solo comprador: los que no tienen OC en SAP son suyos
            con.execute("UPDATE req SET comprador=? WHERE comprador=''", (mayor[0],))
        con.execute("UPDATE req SET cod_excel=id WHERE id IN (SELECT req_id FROM req_eventos "
                    "WHERE texto LIKE 'Importado desde el Excel%')")
    con.commit()


def compradores(con):
    """Compradores conocidos: los de SAP (columna Usuario) y los asignados a requerimientos."""
    nombres = {r[0] for r in con.execute("SELECT DISTINCT usuario FROM sap_oc WHERE usuario<>''")}
    nombres |= {r[0] for r in con.execute("SELECT DISTINCT comprador FROM req WHERE comprador<>''")}
    return sorted(nombres)


class EdicionSimultanea(Exception):
    """Otro usuario (o una carga de SAP) modificó el requerimiento mientras se editaba."""


def config(con, clave, valor=None):
    """Lee (valor=None) o guarda un valor de configuración interna."""
    if valor is None:
        r = con.execute("SELECT valor FROM req_config WHERE clave=?", (clave,)).fetchone()
        return r[0] if r else None
    con.execute("INSERT OR REPLACE INTO req_config VALUES (?,?)", (clave, str(valor)))


def _ahora():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


FORMATOS_FECHA = ("%Y-%m-%d", "%d/%m/%Y", "%Y-%m-%d %H:%M:%S", "%Y-%m-%dT%H:%M:%S", "%d/%m/%y", "%d-%m-%Y", "%d.%m.%Y")


def _iso(v):
    """Fecha (datetime/Timestamp/texto/número de serie de Excel) -> 'YYYY-MM-DD' o '' si no es una fecha."""
    if v is None or (not isinstance(v, str) and pd.isna(v)):
        return ""
    if isinstance(v, str):
        v = v.strip()
        for f in FORMATOS_FECHA:
            try:
                return datetime.strptime(v, f).strftime("%Y-%m-%d")
            except ValueError:
                pass
        return ""
    if isinstance(v, numbers.Number) and not isinstance(v, bool):
        return (datetime(1899, 12, 30) + timedelta(days=int(v))).strftime("%Y-%m-%d") if 20000 < v < 80000 else ""
    return v.strftime("%Y-%m-%d") if hasattr(v, "strftime") else ""


def _norm(v):
    """Valor vacío ('' o None) -> None, para comparar sin confundir 0 con vacío."""
    return None if v is None or v == "" else v


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
    # Proveedores: se suman a los que ya había (cada comprador puede tener su propio Excel con su lista)
    prov = {r[0].upper(): tuple(r) for r in con.execute("SELECT nombre, ruc, origen FROM cat_proveedores")}
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
    with con:      # todo o nada: si algo falla, los catálogos anteriores quedan intactos
        con.execute("DELETE FROM cat_proyectos")
        con.executemany("INSERT OR REPLACE INTO cat_proyectos VALUES (?,?,?,?,?,?,?,?)", filas)
        con.execute("DELETE FROM cat_proveedores")
        con.executemany("INSERT OR REPLACE INTO cat_proveedores VALUES (?,?,?)", list(prov.values()))
    return len(filas), len(prov)


def importar_control(con, ruta_excel, autor="Importación", comprador=""):
    """Importa la tabla CONTROL del Excel como requerimientos de `comprador`. Solo agrega los COD que ese comprador
    aún no tiene en la base. Si algún COD ya lo usa otro requerimiento, las filas nuevas se numeran a continuación
    del mayor COD (el original queda en cod_excel y en el historial).
    Devuelve (nuevos, avisos, renumerados). Todo o nada: si algo falla no queda nada a medias."""
    tmp = _copia_lectura(ruta_excel)
    try:
        c = pd.read_excel(tmp, sheet_name="CONTROL", header=1, dtype=object)
    finally:
        os.remove(tmp)
    c.columns = [_txt(x) for x in c.columns]
    ya = {r[0] for r in con.execute("SELECT cod_excel FROM req WHERE comprador=? AND cod_excel IS NOT NULL", (comprador,))}
    filas, avisos, max_cod = [], [], 0
    for _, r in c.iterrows():
        cod = _int(r.get("COD"))
        if cod is None:
            continue
        max_cod = max(max_cod, cod)      # incluye las filas ya numeradas pero vacías
        if cod in ya or not any(_txt(r.get(h)) for _, h, _ in CAMPOS[1:]):
            continue
        valores = {}
        for campo, enc, tipo in CAMPOS:
            v = r.get(enc)
            valores[campo] = (_iso(v) if tipo == "fecha" else _num(v) if tipo == "num"
                              else _int(v) if tipo == "int" else _txt(v))
            if tipo == "fecha" and not valores[campo] and _txt(v):
                avisos.append(f"COD {cod:03d}: «{enc}» = {_txt(v)!r} no es una fecha válida (quedó vacía)")
        valores["estatus"] = valores["estatus"].upper()
        ya.add(cod)
        filas.append((cod, valores))
    with con:
        usados = {r[0] for r in con.execute("SELECT id FROM req")}
        renumerar = any(cod in usados for cod, _ in filas)
        siguiente = siguiente_id(con)
        for cod, valores in sorted(filas, key=lambda x: x[0]):
            if renumerar:
                valores["id"], siguiente = siguiente, siguiente + 1
            valores.update(comprador=comprador, cod_excel=cod)
            cols = list(valores) + ["creado", "modificado"]
            con.execute(f"INSERT INTO req({','.join(cols)}) VALUES ({','.join('?' * len(cols))})",
                        [valores[k] for k in valores] + [_ahora(), _ahora()])
            evento(con, valores["id"], "Importado desde el Excel de seguimiento"
                   + (f" (COD original en el Excel: {cod:03d})." if renumerar else "."), autor, "sistema")
        if not renumerar:    # los COD se conservaron: los vacíos ya numerados en el Excel no se reutilizan
            config(con, "max_cod_excel", max(max_cod, int(config(con, "max_cod_excel") or 0)))
    return len(filas), avisos, renumerar


# ---------------------------------------------------------------------------
# SAP
# ---------------------------------------------------------------------------
def guardar_snapshot_sap(con, df):
    """Foto del último reporte SAP limpio (procesar_sap.limpiar), para completar requerimientos."""
    cab = df.drop_duplicates("Numero OC")
    con.execute("DELETE FROM sap_oc")
    con.executemany("INSERT INTO sap_oc(numero_oc, fecha, aprobacion, usuario, proveedor, ruc, mon, total, estado, serie) "
                    "VALUES (?,?,?,?,?,?,?,?,?,?)", [
        (int(r["Numero OC"]), _iso(r["Fecha"]), _iso(r["Aprobación"]), r["Usuario"], r["Proveedor"], r["RUC"],
         MON_SAP.get(r["Mon"], r["Mon"]), _num(r["Total Sin IGV"]), r["Estado"], r["Serie"]) for _, r in cab.iterrows()])
    detalle_sap.enlazar(con)               # estado logístico del último reporte de detalle
    con.execute("DELETE FROM sap_fact")
    fx = df[df["N° Factura"] != ""]
    con.executemany("INSERT INTO sap_fact(numero_oc, n_factura, estado, contab, emision, fecha_pago, monto, pendiente) "
                    "VALUES (?,?,?,?,?,?,?,?)", [
        (int(r["Numero OC"]), r["N° Factura"], r["Estado Factura"], _iso(r["Contab. FT"]), _iso(r["Emisión FT"]),
         _iso(r["Fecha Pago"]), _num(r["FT Sin Imp."]), _num(r["Imp. Pendiente"])) for _, r in fx.iterrows()])
    con.commit()


def datos_sap(con, orden):
    """Valores que SAP aporta para una orden (o None si no está en el último reporte)."""
    if not orden:
        return None
    oc = con.execute("SELECT * FROM sap_oc WHERE numero_oc=?", (int(orden),)).fetchone()
    if oc is None:
        return None
    fac = [dict(r) for r in con.execute("SELECT * FROM sap_fact WHERE numero_oc=? ORDER BY contab", (int(orden),))]
    nota = sgm.con_nota_entrada(oc["serie"], oc["estado"], oc["fecha"], oc["ingresado"])
    d = {"emision_oc": oc["fecha"], "aprob_oc": oc["aprobacion"], "proveedor": oc["proveedor"],
         "monto": oc["total"], "moneda": oc["mon"],
         "estado_oc": sgm.estado_visible(oc["serie"], oc["estado"], oc["fecha"], oc["ingresado"]),
         "facturas": fac, "sin_pago": "", "logistica": oc["estado_logistico"] or "", "fecha_entrada": oc["fecha_entrada"] or ""}
    # Bienes recibidos por completo: la fecha de la última nota de entrada es la conformidad (entrega real, OTIF)
    if oc["estado_logistico"] == detalle_sap.ATENDIDA and oc["fecha_entrada"]:
        d["f_conformidad"] = oc["fecha_entrada"]
    if fac:
        d["num_factura"] = ", ".join(f["n_factura"] for f in fac)
        contab = [f["contab"] for f in fac if f["contab"]]
        if contab:
            d["recepcion_fact"] = min(contab)
        # Pagada: estado Pagado, o un saldo menor que es solo diferencia de montos
        saldada = lambda f: f["estado"] == "Pagado" or (f.get("pendiente") is not None and f["pendiente"] <= SALDO_MENOR)
        pagadas = [f for f in fac if saldada(f) and f["fecha_pago"]]
        facturado = sum(f["monto"] or 0 for f in fac)
        # «Cerrado» con nota de entrada solo dice que se recibió todo: para estar completa debe estar facturada
        completa = oc["estado"] == "Cerrado" and not nota or bool(oc["total"]) and (
            facturado >= 0.995 * oc["total"] or oc["total"] - facturado <= SALDO_MENOR)
        # Fecha de pago: todas las facturas pagadas Y la orden cerrada o facturada al 100 % (si no, es un pago parcial)
        if len(pagadas) == len(fac) and completa:
            d["fecha_pago"] = max(f["fecha_pago"] for f in pagadas)
        elif len(pagadas) == len(fac):
            d["sin_pago"] = "pago parcial: la OC sigue en proceso y no está facturada al 100 %"
        else:
            d["sin_pago"] = "hay facturas de la OC pendientes de pago"
    return d


def usos_oc(con):
    """{N° de orden: [ids de requerimientos que la usan]}"""
    usos = {}
    for rid, orden in con.execute("SELECT id, orden FROM req WHERE orden IS NOT NULL ORDER BY id"):
        usos.setdefault(orden, []).append(rid)
    return usos


# Campos que aporta SAP. Los que vinieron de SAP se mantienen al día; los escritos por el usuario no se tocan;
# los vacíos se completan (salvo que el usuario los haya vaciado a propósito).
CAMPOS_SAP = ["emision_oc", "aprob_oc", "proveedor", "monto", "moneda", "num_factura", "recepcion_fact", "fecha_pago",
              "f_conformidad"]
# Vinculados: si la orden está en SAP, SAP manda siempre (lo aprobado en SAP es lo correcto).
# Excepción: el monto de una OC que comparten varios requerimientos (el total SAP es de la orden completa).
CAMPOS_VINCULADOS = ["monto", "moneda"]
# Saldo (en la moneda de la OC / factura) que se considera diferencia de montos y no un pendiente real
SALDO_MENOR = 10.0


def _distinto(campo, a, b):
    a, b = _norm(a), _norm(b)
    if campo == "monto":
        return (a is None) != (b is None) or (a is not None and abs(float(a) - float(b)) > 0.005)
    return a != b


def completar_desde_sap(con, req_id=None, autor="Sistema"):
    """Sincroniza los requerimientos con su orden en el último reporte SAP. Devuelve cuántos cambiaron."""
    q = "SELECT * FROM req WHERE orden IS NOT NULL" + (" AND id=?" if req_id else "")
    usos = usos_oc(con)
    cambiados = 0
    for r in con.execute(q, (req_id,) if req_id else ()).fetchall():
        d = datos_sap(con, r["orden"])
        if not d:
            continue
        desde = set(json.loads(r["desde_sap"] or "[]"))
        manual = set(json.loads(r["manual"] or "[]"))
        compartida = len(usos.get(r["orden"], [])) > 1
        llenados, actualizados = {}, {}
        for campo in CAMPOS_SAP:
            v, actual = _norm(d.get(campo)), _norm(r[campo])
            if campo == "monto" and compartida:
                desde.discard("monto")       # se respeta el reparto del usuario entre los requerimientos
                continue
            if campo in CAMPOS_VINCULADOS or campo in desde:
                if v is None:
                    if campo in desde and actual is not None:      # SAP ya no tiene el dato: se limpia
                        actualizados[campo] = None
                        desde.discard(campo)
                    continue
                if _distinto(campo, actual, v):
                    (llenados if actual is None else actualizados)[campo] = v
                desde.add(campo)
            elif actual is None and v is not None and campo not in manual:
                llenados[campo] = v
                desde.add(campo)
        cambios = {**llenados, **actualizados}
        if cambios or sorted(desde) != json.loads(r["desde_sap"] or "[]"):
            sets = "".join(f"{k}=?, " for k in cambios)
            vals = [("" if v is None and k != "monto" else v) for k, v in cambios.items()]
            con.execute(f"UPDATE req SET {sets}desde_sap=?, modificado=? WHERE id=?",
                        [*vals, json.dumps(sorted(desde)), _ahora(), r["id"]])
        ver = lambda k, v: "(vacío)" if v in (None, "") else fmt(v) if k in CAMPOS_FECHA or k == "f_conformidad" else (
            f"{v:,.2f}" if k == "monto" else v)
        if llenados:
            det = "; ".join(f"{ETIQUETAS[k]}: {ver(k, v)}" for k, v in llenados.items())
            evento(con, r["id"], f"Completado desde SAP (OC {r['orden']}): {det}", autor, "sistema")
        if actualizados:
            det = "; ".join(f"{ETIQUETAS[k]}: {ver(k, r[k])} → {ver(k, v)}" for k, v in actualizados.items())
            if "fecha_pago" in actualizados and actualizados["fecha_pago"] is None and d["sin_pago"]:
                det += f" ({d['sin_pago']})"
            evento(con, r["id"], f"Actualizado según SAP (OC {r['orden']}): {det}", autor, "sistema")
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
    """Siguiente COD libre: después del mayor de la base y del mayor numerado en el Excel importado."""
    maximo = con.execute("SELECT MAX(id) FROM req").fetchone()[0] or 0
    return max(maximo, int(config(con, "max_cod_excel") or 0)) + 1


def guardar(con, valores, autor):
    """Crea o actualiza un requerimiento; registra en el historial cada campo que cambia.
    valores['_sap']: campos que el formulario completó con «Traer de SAP» (quedan marcados como de SAP)."""
    valores = dict(valores)
    sap_form = set(valores.pop("_sap", ()))
    leido = valores.pop("_modificado", None)
    rid = valores.get("id")
    prev = con.execute("SELECT * FROM req WHERE id=?", (rid,)).fetchone() if rid else None
    if prev is not None and leido is not None and prev["modificado"] != leido:
        ult = con.execute("SELECT autor, fecha FROM req_eventos WHERE req_id=? ORDER BY fecha DESC, id DESC LIMIT 1",
                          (rid,)).fetchone()
        raise EdicionSimultanea(f"{ult['autor']} ({fmt(ult['fecha'][:10])} {ult['fecha'][11:16]})" if ult else "otro usuario")
    campos = [c for c, _, _ in CAMPOS if c != "id"] + ["notas", "comprador"] + CAMPOS_KPI
    if "comprador" not in valores:
        valores["comprador"] = prev["comprador"] if prev is not None else ""
    if prev is None:
        cols = ["id"] + campos + ["desde_sap", "creado", "modificado"]
        for intento in range(5):
            nuevo_id = rid or siguiente_id(con)
            try:
                con.execute(f"INSERT INTO req({','.join(cols)}) VALUES ({','.join('?' * len(cols))})",
                            [nuevo_id] + [valores.get(c) for c in campos]
                            + [json.dumps(sorted(sap_form & set(CAMPOS_SAP))), _ahora(), _ahora()])
                break
            except sqlite3.IntegrityError:
                if rid or intento == 4:      # el otro comprador acaba de usar ese COD: se toma el siguiente
                    raise
        rid = nuevo_id
        evento(con, rid, "Requerimiento registrado.", autor, "cambio")
    else:
        desde = set(json.loads(prev["desde_sap"] or "[]"))
        manual = set(json.loads(prev["manual"] or "[]"))
        cambios = []
        # El formulario muestra el monto con 2 decimales: si no se tocó, se conserva el valor original
        if valores.get("monto") is not None and prev["monto"] is not None and not _distinto("monto", valores["monto"], prev["monto"]):
            valores["monto"] = prev["monto"]
        # Si cambia la orden, los datos que vinieron de SAP eran de la orden anterior: se limpian y se traen los nuevos
        limpiados = set()
        if _norm(valores.get("orden")) != _norm(prev["orden"]):
            for c in desde - sap_form:
                if not _distinto(c, valores.get(c), prev[c]):
                    valores[c] = None if c == "monto" else ""
                    limpiados.add(c)
            desde -= limpiados
        for c in campos:
            nuevo, viejo = valores.get(c), prev[c]
            if _distinto(c, nuevo, viejo):
                ver = lambda v: "(vacío)" if _norm(v) is None else (
                    fmt(v) if c in CAMPOS_FECHA or c == "f_conformidad" else f"{v:,.2f}" if c in ("monto", "monto_ref") else v)
                if c != "notas" and c not in limpiados:
                    cambios.append(f"{ETIQUETAS[c]}: {ver(viejo)} → {ver(nuevo)}")
                if c in sap_form:
                    desde.add(c)
                    manual.discard(c)
                elif c not in limpiados:
                    desde.discard(c)          # editado a mano: ya no es dato de SAP
                    if c in CAMPOS_SAP:
                        # vaciado a propósito: SAP no lo vuelve a llenar; con valor: es del usuario
                        (manual.add if _norm(nuevo) is None else manual.discard)(c)
        if limpiados:
            cambios.append("se quitaron los datos SAP de la orden anterior ("
                           + ", ".join(ETIQUETAS[c] for c in CAMPOS_SAP if c in limpiados) + ")")
        if not cambios and _norm(valores.get("notas")) == _norm(prev["notas"]):
            return rid
        sets = ", ".join(f"{c}=?" for c in campos)
        con.execute(f"UPDATE req SET {sets}, desde_sap=?, manual=?, modificado=? WHERE id=?",
                    [valores.get(c) for c in campos] + [json.dumps(sorted(desde)), json.dumps(sorted(manual)), _ahora(), rid])
        if cambios:
            evento(con, rid, "; ".join(cambios), autor, "cambio")
    con.commit()
    if valores.get("orden"):
        completar_desde_sap(con, rid)
        if prev is None or _norm(valores.get("orden")) != _norm(prev["orden"]):
            completar_desde_sap(con)      # otras que comparten la orden (el monto pasa a ser por requerimiento)
    return rid


def etapa(r):
    """(nombre de etapa, fecha ISO en que empezó). Usa r['sap_estado'] y r['sap_n_fact'] si existen."""
    cierre = ESTATUS_CIERRE.get((r["estatus"] or "").upper())
    if r["fecha_pago"]:
        return "Pagado", r["fecha_pago"]
    if cierre:
        return cierre, r["modificado"][:10] if r.get("modificado") else ""
    if r.get("sap_estado") in SAP_CIERRE and not r.get("sap_nota_entrada") and not r.get("sap_n_fact") and not r["num_factura"]:
        return SAP_CIERRE[r["sap_estado"]], ""
    for nombre, campo in reversed(ETAPAS):
        if r[campo]:
            return nombre, r[campo]
        if campo == "emision_oc" and r["orden"]:
            # Tiene N° de orden aunque no se haya registrado la fecha de emisión
            return nombre, r["aprob_valo"] or r["envio_aprob_valo"] or r["envio_valo"] or r["f_solic"]
    return "Por valorizar", r["f_solic"]


def listar(con):
    proys = {r["codigo"]: dict(r) for r in con.execute("SELECT * FROM cat_proyectos")}
    sap = {r["numero_oc"]: r["estado"] for r in con.execute("SELECT numero_oc, estado FROM sap_oc")}
    nota_entrada = {r["numero_oc"] for r in con.execute("SELECT numero_oc, serie, estado, fecha, ingresado FROM sap_oc")
                    if sgm.con_nota_entrada(r["serie"], r["estado"], r["fecha"], r["ingresado"])}
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
        d["sap_nota_entrada"] = d["orden"] in nota_entrada
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


STATUS_RESUMEN = ["Completado", "Pendiente de pago", "Facturación parcial", "Pendiente de factura", "Sin OC (en trámite)",
                  "OC no está en SAP", "Anulada"]


def estado_pago(con, orden):
    """Status de pago de una OC según SAP: (status, detalle, cifras).
    «Completado» = nada pendiente de pago: facturas pagadas (o saldo menor por diferencia de montos), o la OC se
    cerró sin factura (salvo que tenga nota de entrada: ahí falta la factura)."""
    oc = con.execute("SELECT * FROM sap_oc WHERE numero_oc=?", (int(orden),)).fetchone()
    if oc is None:
        return "OC no está en SAP", "La orden no aparece en el último reporte «Control de órdenes»", {}
    fac = [dict(r) for r in con.execute(
        "SELECT * FROM facturas WHERE numero_oc=? AND situacion='Vigente' ORDER BY emision", (int(orden),))]
    mon, total = MON_SAP.get(oc["mon"], oc["mon"]), oc["total"] or 0
    facturado = sum(f["ft_sin_imp"] or 0 for f in fac)
    saldo = sum(max(f["imp_pendiente"] or 0, 0) for f in fac if f["estado_factura"] != "Pagado")
    pagos = [f["fecha_pago"] for f in fac if f["fecha_pago"]]
    cifras = {"moneda": mon, "total": total, "facturado": facturado, "pct": facturado / total if total else None,
              "saldo": saldo, "ultimo_pago": max(pagos) if pagos else "", "estado_oc": oc["estado"], "n_fact": len(fac),
              "proveedor": oc["proveedor"], "fecha_oc": oc["fecha"]}
    nota = sgm.con_nota_entrada(oc["serie"], oc["estado"], oc["fecha"], oc["ingresado"])
    if oc["estado"] == "Cancelado":
        return "Anulada", "OC cancelada en SAP", cifras
    if saldo > SALDO_MENOR:
        hoy = date.today().isoformat()
        pend = [f for f in fac if f["estado_factura"] != "Pagado" and (f["imp_pendiente"] or 0) > SALDO_MENOR]
        venc = [f for f in pend if f["vencimiento"] and f["vencimiento"] < hoy]
        prox = min((f["vencimiento"] for f in pend if f["vencimiento"] and f["vencimiento"] >= hoy), default="")
        det = f"{len(pend)} factura{'s' if len(pend) > 1 else ''} por pagar: {mon} {saldo:,.2f}"
        det += f" · {len(venc)} vencida{'s' if len(venc) > 1 else ''}" if venc else (f" · vence {fmt(prox)}" if prox else "")
        return "Pendiente de pago", det, cifras
    if not fac:
        if oc["estado"] == "Cerrado" and not nota:
            return "Completado", "OC cerrada sin factura: no hay nada por pagar", cifras
        return "Pendiente de factura", ("Material ingresado (nota de entrada); falta la factura" if nota
                                        else "OC en proceso, aún sin factura"), cifras
    por_facturar = total - facturado
    if (oc["estado"] != "Cerrado" or nota) and por_facturar > SALDO_MENOR and facturado < 0.995 * total:
        return "Facturación parcial", (f"Facturado {facturado / total:.0%}; falta facturar {mon} {por_facturar:,.2f} "
                                       f"(lo facturado está pagado)"), cifras
    det = (f"{len(fac)} factura{'s' if len(fac) > 1 else ''} pagada{'s' if len(fac) > 1 else ''}"
           + (f" · último pago {fmt(max(pagos))}" if pagos else ""))
    if saldo > 0.005:
        det += f" · saldo menor {mon} {saldo:,.2f} (diferencia de montos)"
    if oc["estado"] == "Cerrado" and abs(por_facturar) > 0.005 and facturado < 0.995 * total:
        det += f" · OC cerrada con {mon} {por_facturar:,.2f} sin facturar"
    return "Completado", det, cifras


# Tipos que no entran en el resumen (los vuelos no pasan por OC y su seguimiento no es relevante aquí)
TIPOS_FUERA_RESUMEN = ("VUELO",)


def resumen_sap(con, filas):
    """Filas del resumen (una por requerimiento, sin vuelos) con lo que dice SAP y su status de pago."""
    salida = []
    for f in filas:
        if (f["tipo"] or "").upper() in TIPOS_FUERA_RESUMEN:
            continue
        if not f["orden"] and f["etapa"] == "Anulado":
            st, det, c = "Anulada", "Anulado en el control (sin OC)", {}
        elif not f["orden"] and f["etapa"] in CERRADAS:
            st, det, c = "Completado", f"Sin OC en SAP ({f['tipo'] or 'sin tipo'}); cerrado en el control: {f['etapa']}", {}
        elif not f["orden"]:
            st, det, c = "Sin OC (en trámite)", f"Etapa: {f['etapa']}", {}
        else:
            st, det, c = estado_pago(con, f["orden"])
        salida.append({**c, "id": f["id"], "proyecto": f["proyecto"], "orden": f["orden"], "status": st, "detalle": det,
                       "proveedor": c.get("proveedor") or f["proveedor"], "moneda": c.get("moneda") or f["moneda"],
                       "total": c.get("total", f["monto"])})
    return salida


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
COLORES_STATUS = {"Completado": "#C6EFCE", "Pendiente de pago": "#FFC7CE", "Facturación parcial": "#FFEB9C",
                  "Pendiente de factura": "#FFEB9C", "Sin OC (en trámite)": "#DDEBF7", "OC no está en SAP": "#EDEDED",
                  "Anulada": "#EDEDED"}


def _hoja_resumen(wb, F, hdr, titulo, sub, filas):
    """Hoja RESUMEN: una fila por requerimiento con solo datos de SAP y el status de pago."""
    ws = wb.add_worksheet("RESUMEN")
    ws.hide_gridlines(2)
    ws.write("A1", "Resumen de requerimientos según SAP", titulo)
    ws.write("A2", f"«Completado» = sin pendientes de pago: facturas pagadas (saldos de hasta {SALDO_MENOR:,.0f} se toman "
                   f"como diferencia de montos) u OC cerrada sin factura. DÍAS = antigüedad de la OC en lo pendiente. "
                   f"No incluye vuelos. Datos del último reporte SAP cargado.", sub)
    cuenta = {}
    for f in filas:
        cuenta[f["status"]] = cuenta.get(f["status"], 0) + 1
    saldos = {}
    for f in filas:
        if f["status"] == "Pendiente de pago":
            saldos[f["moneda"]] = saldos.get(f["moneda"], 0) + f["saldo"]
    ws.write("A3", "  ·  ".join(f"{s}: {cuenta[s]}" for s in STATUS_RESUMEN if s in cuenta)
             + ("   |   Por pagar: " + "  ·  ".join(f"{m} {v:,.2f}" for m, v in sorted(saldos.items())) if saldos else ""),
             F(bold=True, font_color="#1F3864"))
    cols = ["COD", "PROYECTO", "PROVEEDOR", "ORDEN", "FECHA OC", "DÍAS", "ESTADO OC", "MONEDA", "MONTO OC", "FACTURADO",
            "% FACTURADO", "SALDO POR PAGAR", "ÚLTIMO PAGO", "STATUS", "DETALLE"]
    orden = {s: i for i, s in enumerate(STATUS_RESUMEN)}
    filas = sorted(filas, key=lambda f: (orden.get(f["status"], 9), -(f["saldo"] if f.get("saldo") else 0), f["id"]))
    fecha = lambda iso: datetime.strptime(iso, "%Y-%m-%d") if iso else None
    dias = lambda f: ((date.today() - fecha(f["fecha_oc"]).date()).days
                      if f.get("fecha_oc") and f["status"] not in ("Completado", "Anulada") else None)
    datos = [[f["id"], f["proyecto"] or None, f["proveedor"] or None, f["orden"], fecha(f.get("fecha_oc")), dias(f),
              f.get("estado_oc") or None, f["moneda"] or None, f["total"], f.get("facturado"), f.get("pct"),
              f.get("saldo") if f.get("saldo") else None, fecha(f.get("ultimo_pago")), f["status"], f["detalle"]]
             for f in filas]
    fd, fm, fp = F(num_format="dd/mm/yyyy"), F(num_format="#,##0.00"), F(num_format="0%")
    formatos = {"FECHA OC": fd, "ÚLTIMO PAGO": fd, "MONTO OC": fm, "FACTURADO": fm, "SALDO POR PAGAR": fm, "% FACTURADO": fp}
    ws.add_table(4, 0, 4 + max(len(datos), 1), len(cols) - 1, {
        "name": "Resumen", "data": datos, "style": "Table Style Light 9",
        "columns": [{"header": c, "header_format": hdr, **({"format": formatos[c]} if c in formatos else {})} for c in cols]})
    anchos = [6, 11, 38, 11, 11, 6, 11, 8, 13, 13, 10, 14, 11, 20, 62]
    for i, w in enumerate(anchos):
        ws.set_column(i, i, w)
    col_st = cols.index("STATUS")
    for st, color in COLORES_STATUS.items():
        ws.conditional_format(5, col_st, 4 + max(len(datos), 1), col_st,
                              {"type": "cell", "criteria": "==", "value": f'"{st}"', "format": F(bg_color=color, bold=True)})
    ws.freeze_panes(5, 1)


def exportar_excel(con, salida, plazos=None, comprador=None):
    """Genera el Excel de seguimiento. Si `salida` está abierto en Excel, guarda una copia con la hora.
    `plazos`: {id: 'Atrasado' | 'Atención' | 'En plazo'} del semáforo (columna PLAZO).
    `comprador`: solo los requerimientos de ese comprador (None = todos). Devuelve la ruta escrita."""
    plazos = plazos or {}
    import xlsxwriter
    filas = [f for f in listar(con) if comprador is None or f["comprador"] == comprador]
    destino, salida = salida, os.path.join(os.path.dirname(salida), f"~tmp_req_{os.getpid()}.xlsx")
    # strings_to_formulas=False: un texto que empiece con «=» (nota, proveedor…) se guarda como texto
    wb = xlsxwriter.Workbook(salida, {"default_format_properties": {"font_name": "Arial", "font_size": 10},
                                      "strings_to_formulas": False})
    F = lambda **k: wb.add_format({"font_name": "Arial", "font_size": 10, **k})
    AZUL = "#1F3864"
    hdr = F(bold=True, font_color="white", bg_color=AZUL, text_wrap=True, valign="vcenter")
    grp = F(bold=True, font_color="white", bg_color="#2F5597", align="center")
    fd, fm, fsap = F(num_format="dd/mm/yyyy"), F(num_format="#,##0.00"), F(num_format="dd/mm/yyyy", font_color="#2F5597", italic=True)
    titulo, sub = F(bold=True, font_size=14), F(font_color="#595959")
    _hoja_resumen(wb, F, hdr, titulo, sub, resumen_sap(con, filas))

    # --- CONTROL (mismas columnas que el Excel original + etapa y tiempos)
    ws = wb.add_worksheet("CONTROL")
    extra = ["COMPRADOR", "CONTROLLER", "CONTROLLER2", "JEFE DE PROYECTO", "ETAPA", "DÍAS EN ETAPA", "PLAZO", "ESTADO OC SAP",
             "ÚLTIMO COMENTARIO OC", "DATOS DESDE SAP", "MONTO REFERENCIAL", "TIPO DE REFERENCIA",
             "FECHA DE CONFORMIDAD", "NOTAS"]
    tramos = [t[0] for t in TRAMOS]
    enc = [h for _, h, _ in CAMPOS] + extra + [f"DÍAS: {t}" for t in tramos]
    pos = {h: i for i, h in enumerate(enc)}
    grupos = [("DETALLE DE REQUERIMIENTO", "FECHA INICIO", "MONEDA"), ("FLUJO VALORIZACION", "ENVIO DE VALORIZACION", "FECHA DE APROBACION VALO"),
              ("FLUJO ORDEN", "EMISION DE ORDEN", "ORDEN"), ("FLUJO PAGO", "RECEPCION DE FACTURA", "FECHA DE PAGO"),
              ("SEGUIMIENTO", "COMPRADOR", "NOTAS"), ("TIEMPOS (días)", enc[len(CAMPOS) + len(extra)], enc[-1])]
    for g, a, b in grupos:
        ws.merge_range(0, pos[a], 0, pos[b], g, grp)
    n = len(filas)
    datos = []
    for f in filas:
        fila = []
        for campo, _, tipo in CAMPOS:
            v = f[campo]
            fila.append(datetime.strptime(v, "%Y-%m-%d") if tipo == "fecha" and v else (v if v not in ("", None) else None))
        fila += [f["comprador"] or None, f["controller"], f["controller2"], f["jefe"], f["etapa"], f["dias_etapa"], plazos.get(f["id"]) or None,
                 f["sap_estado"] or None,
                 f["com_oc"] or None, ", ".join(ETIQUETAS[c] for c in f["desde_sap"]), f["monto_ref"],
                 f["tipo_ref"] or None, datetime.strptime(f["f_conformidad"], "%Y-%m-%d") if f["f_conformidad"] else None,
                 f["notas"] or None]
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
    for h in ("COMPRADOR", "CONTROLLER", "CONTROLLER2", "JEFE DE PROYECTO"):
        ws.set_column(pos[h], pos[h], 26)
    ws.set_column(pos["ETAPA"], pos["ETAPA"], 24)
    ws.set_column(pos["DATOS DESDE SAP"], pos["DATOS DESDE SAP"], 30)
    ws.set_column(pos["ESTADO OC SAP"], pos["ESTADO OC SAP"], 13)
    ws.set_column(pos["ÚLTIMO COMENTARIO OC"], pos["ÚLTIMO COMENTARIO OC"], 40)
    ws.set_column(pos["NOTAS"], pos["NOTAS"], 40)
    ws.set_column(pos["MONTO REFERENCIAL"], pos["MONTO REFERENCIAL"], 14, fm)
    ws.set_column(pos["TIPO DE REFERENCIA"], pos["TIPO DE REFERENCIA"], 18)
    ws.set_column(pos["FECHA DE CONFORMIDAD"], pos["FECHA DE CONFORMIDAD"], 13, fd)
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
    pe.write_row(3, 0, ["Etapa", "N°", "Atrasados", "En atención", "Días promedio en la etapa", "Máx. días",
                        "Monto PEN", "Monto USD"], hdr)
    rojo = F(num_format='0;-0;"-"', bg_color="#F8CBAD")
    amarillo = F(num_format='0;-0;"-"', bg_color="#FFE699")
    for i, nombre in enumerate(NOMBRES_ETAPA[:-1]):
        r = 4 + i
        pe.write(r, 0, nombre, F(bold=nombre in ETAPAS_HASTA_OC))
        pe.write_formula(r, 1, f'=COUNTIFS(Tabla1[ETAPA],$A{r + 1})')
        pe.write_formula(r, 2, f'=COUNTIFS(Tabla1[ETAPA],$A{r + 1},Tabla1[PLAZO],"Atrasado")', rojo)
        pe.write_formula(r, 3, f'=COUNTIFS(Tabla1[ETAPA],$A{r + 1},Tabla1[PLAZO],"Atención")', amarillo)
        pe.write_formula(r, 4, f'=IFERROR(AVERAGEIFS(Tabla1[DÍAS EN ETAPA],Tabla1[ETAPA],$A{r + 1}),"-")', F(num_format="0.0"))
        pe.write_formula(r, 5, f'=IF(B{r + 1}=0,"-",_xlfn.MAXIFS(Tabla1[DÍAS EN ETAPA],Tabla1[ETAPA],$A{r + 1}))')
        pe.write_formula(r, 6, f'=SUMIFS(Tabla1[MONTO],Tabla1[ETAPA],$A{r + 1},Tabla1[MONEDA],"PEN")', F(num_format='#,##0.00;-#,##0.00;"-"'))
        pe.write_formula(r, 7, f'=SUMIFS(Tabla1[MONTO],Tabla1[ETAPA],$A{r + 1},Tabla1[MONEDA],"USD")', F(num_format='#,##0.00;-#,##0.00;"-"'))
    fin = 4 + len(NOMBRES_ETAPA) - 1
    total = F(bold=True, bg_color="#F2F2F2")
    pe.write(fin, 0, "Total pendientes hasta envío de OC", total)
    for c, col in enumerate("BCD", start=1):
        pe.write_formula(fin, c, f"=SUM({col}5:{col}{4 + len(ETAPAS_HASTA_OC)})", total)
    pe.write(fin + 2, 0, "Plazo: «Atrasado» y «Atención» según los umbrales del semáforo de la aplicación (⚙ Plazos…).", sub)
    pe.set_column("A:A", 34)
    pe.set_column("B:H", 15)

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
    try:
        os.replace(salida, destino)
    except PermissionError:          # el archivo del día está abierto en Excel
        destino = destino.replace(".xlsx", f" ({datetime.now():%H%M%S}).xlsx")
        os.replace(salida, destino)
    return destino
