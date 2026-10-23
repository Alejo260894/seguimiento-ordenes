"""
Reporte de detalle «Seguimiento OC → Entrada → Factura → Pago» (consulta SAP B1 / HANA, exportada a .xlsx).

Complementa al «Control de órdenes» (que sigue siendo la fuente general):
- Ingreso de material (nota de entrada) por línea de OC: estado logístico real de cada OC.
- Facturas con su N° interno (clave única) y el estado «Pago parcial».
- Líneas de la OC con proyecto, unidad de negocio y área.
Solo se guardan las OC aprobadas (con fecha de aprobación), igual que en el Control.

Granularidad del archivo: una fila = línea de OC × línea de entrada × línea de factura; por eso cada nivel se
deduplica por su propia clave antes de guardarlo.

Atención de las OC de bienes (el registro de notas de entrada es reciente, así que la factura también cuenta):
- Línea: «Ingresada» (nota de entrada por toda la cantidad), «Atendida por factura» (facturada al 100 % sin nota),
  «Ingreso parcial» o «Pendiente de ingreso» (sin ingreso ni factura).
- OC: «Atendida» (todas sus líneas atendidas, o la OC facturada por completo), «Atención parcial» o «Pendiente de
  atención». Si SAP ya la cerró: «Cerrada con atención parcial» o «Cerrada sin ingreso».
- Las OC de servicios no llevan nota de entrada: se siguen por factura («Servicio»).
"""
import os
import re
import shutil
import tempfile
import zipfile
from datetime import datetime

import pandas as pd

COLUMNAS = ["N° OC", "Estado OC", "Serie", "Fecha OC", "Aprobación", "Tipo OC", "Línea", "Código", "Descripción",
            "Cant. OC", "UM", "Total Línea Sin IGV", "Proyecto", "Unidad Negocio", "Área", "N° Entrada",
            "Fecha Entrada", "Cant. Ingresada", "Almacén", "Total Ingresado Línea", "Pendiente Ingreso",
            "Estado Logístico", "N° Factura", "Tipo", "N° Interno FT", "Emisión FT", "Contab. FT", "Venc. FT",
            "Estado Factura", "Fecha Pago", "Imp. Pendiente", "RUC", "Moneda"]

# Estado logístico de la OC (bienes), del más pendiente al más avanzado
SIN_DETALLE = "(sin detalle)"
SERVICIO = "Servicio"
PENDIENTE, PARCIAL, ATENDIDA = "Pendiente de atención", "Atención parcial", "Atendida"
LOGISTICA = [PENDIENTE, PARCIAL, ATENDIDA, "Cerrada con atención parcial", "Cerrada sin ingreso", "OC cancelada", SERVICIO]
PENDIENTES = (PENDIENTE, PARCIAL)                     # OC abiertas con material por ingresar
LINEA_PENDIENTE, LINEA_PARCIAL = "Pendiente de ingreso", "Ingreso parcial"
LINEAS_ATENDIDAS = ("Ingresada", "Atendida por factura")
TOLERANCIA = 0.005


def es_detalle(ruta):
    """¿El archivo es el reporte de detalle? (lee solo el encabezado)."""
    if not ruta.lower().endswith((".xlsx", ".xlsm")):
        return False
    try:
        cab = pd.read_excel(ruta, engine="calamine", nrows=0)
    except Exception:  # noqa: BLE001 — no es un Excel legible: que lo trate el lector del Control
        return False
    return "Estado Logístico" in cab.columns and "N° Interno FT" in cab.columns


def _limpiar_xlsx(ruta):
    """Copia del .xlsx sin caracteres de control dentro de los XML (SAP a veces los exporta y rompen la lectura)."""
    tmp = os.path.join(tempfile.gettempdir(), f"detalle_{os.getpid()}.xlsx")
    malos = re.compile(rb"[\x00-\x08\x0b\x0c\x0e-\x1f]")
    with zipfile.ZipFile(ruta) as z, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as out:
        for item in z.infolist():
            datos = z.read(item.filename)
            if item.filename.endswith(".xml"):
                datos = malos.sub(b"", datos)
            out.writestr(item, datos)
    return tmp


def leer(ruta):
    """DataFrame del reporte (todo como texto). ValueError con un mensaje claro si no es el archivo esperado."""
    if ruta.lower().endswith(".csv"):
        raise ValueError("Del reporte de detalle usa el Excel (.xlsx): el CSV de SAP sale en UTF-16 y sin comillas, "
                         "y se desordenan las columnas.")
    copia = os.path.join(tempfile.gettempdir(), f"detalle_copia_{os.getpid()}.xlsx")   # puede estar abierto en Excel
    shutil.copyfile(ruta, copia)
    try:
        tmp = _limpiar_xlsx(copia)
    finally:
        os.remove(copia)
    try:
        df = pd.read_excel(tmp, engine="calamine", dtype=str).fillna("")
    finally:
        os.remove(tmp)
    faltan = [c for c in COLUMNAS if c not in df.columns]
    if faltan:
        raise ValueError("El archivo no parece el reporte de detalle «Seguimiento OC → Entrada → Factura → Pago»: "
                         f"faltan las columnas {', '.join(faltan[:6])}.")
    for c in df.columns:
        df[c] = df[c].str.strip()
    return df


def _iso(s):
    s = (s or "").strip()
    if not s:
        return ""
    for f in ("%d/%m/%Y", "%Y-%m-%d", "%Y-%m-%d %H:%M:%S"):
        try:
            return datetime.strptime(s, f).strftime("%Y-%m-%d")
        except ValueError:
            pass
    return ""


def _num(s):
    try:
        return float(str(s).replace(",", ""))
    except ValueError:
        return 0.0


def _oc(s):
    try:
        return int(float(s))
    except ValueError:
        return None


def estado_linea(cantidad, ingresado, total, facturado):
    """(estado, fracción atendida) de una línea de OC de bienes: la nota de entrada o la factura la atienden."""
    rec = min(ingresado / cantidad, 1) if cantidad else (1.0 if ingresado > 0 else 0.0)
    fac = min(facturado / total, 1) if total else 0.0
    if rec >= 1 - TOLERANCIA:
        return "Ingresada", 1.0
    if fac >= 1 - TOLERANCIA:
        return "Atendida por factura", 1.0
    if rec > 0 or fac > 0:
        return LINEA_PARCIAL, max(rec, fac)
    return LINEA_PENDIENTE, 0.0


def estado_oc(tipo_oc, estado, lineas):
    """(estado logístico, fracción atendida) de una OC; lineas = [(total_linea, estado_linea, fraccion, facturado)]."""
    if estado == "Cancelado":
        return "OC cancelada", None
    if tipo_oc == SERVICIO:
        return SERVICIO, None
    total = sum(t for t, _, _, _ in lineas)
    facturado = sum(f for _, _, _, f in lineas)
    avance = sum(t * fr for t, _, fr, _ in lineas) / total if total else (
        sum(fr for _, _, fr, _ in lineas) / len(lineas) if lineas else 0)
    if (total and facturado >= (1 - TOLERANCIA) * total) or all(e in LINEAS_ATENDIDAS for _, e, _, _ in lineas):
        return ATENDIDA, 1.0
    parcial = any(e != LINEA_PENDIENTE for _, e, _, _ in lineas)
    if estado == "Cerrado":
        return ("Cerrada con atención parcial" if parcial else "Cerrada sin ingreso"), avance
    return (PARCIAL if parcial else PENDIENTE), avance


def guardar(con, df, archivo):
    """Reemplaza el detalle guardado con el del reporte (solo OC aprobadas). Devuelve un resumen."""
    df = df[df["Aprobación"] != ""].copy()
    df["oc"] = df["N° OC"].map(_oc)
    df = df[df["oc"].notna()]
    lineas = df.drop_duplicates(["oc", "Línea"])
    entradas = df[df["N° Entrada"] != ""].drop_duplicates(["oc", "Línea", "N° Entrada"])
    facturas = df[df["N° Interno FT"] != ""].drop_duplicates(["N° Interno FT", "oc"])
    fecha_ent = entradas.groupby(["oc", "Línea"])["Fecha Entrada"].agg(lambda s: max(_iso(x) for x in s))
    # Monto facturado por línea (facturas no anuladas; cada línea de factura una sola vez)
    ft = df[(df["N° Interno FT"] != "") & (df["Estado Factura"] != "Anulado")].drop_duplicates(["oc", "Línea", "N° Interno FT"])
    fact_lin = ft.groupby(["oc", "Línea"])["FT Sin Imp. (línea)"].agg(lambda s: sum(_num(x) for x in s))
    est_lin = {}
    for _, r in lineas.iterrows():
        f = fact_lin.get((r["oc"], r["Línea"]), 0.0)
        e, fr = estado_linea(_num(r["Cant. OC"]), _num(r["Total Ingresado Línea"]), _num(r["Total Línea Sin IGV"]), f)
        est_lin[(r["oc"], r["Línea"])] = (e, fr, f) if r["Tipo OC"] != SERVICIO else (SERVICIO, None, f)
    _asegurar_columnas(con)

    with con:
        for t in ("detalle_lineas", "detalle_entradas", "detalle_facturas", "detalle_oc"):
            con.execute(f"DELETE FROM {t}")
        con.executemany("""INSERT INTO detalle_lineas(numero_oc, linea, tipo_oc, codigo, descripcion, cantidad, um,
                           total_linea, moneda, proyecto, unidad_negocio, area, ingresado, pendiente_ingreso,
                           estado_logistico, fecha_entrada, facturado, estado_linea) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""", [
            (int(r["oc"]), int(_num(r["Línea"])), r["Tipo OC"], r["Código"], r["Descripción"], _num(r["Cant. OC"]),
             r["UM"], _num(r["Total Línea Sin IGV"]), r["Moneda"], r["Proyecto"], r["Unidad Negocio"], r["Área"],
             _num(r["Total Ingresado Línea"]), _num(r["Pendiente Ingreso"]), r["Estado Logístico"],
             fecha_ent.get((r["oc"], r["Línea"]), ""), est_lin[(r["oc"], r["Línea"])][2],
             est_lin[(r["oc"], r["Línea"])][0]) for _, r in lineas.iterrows()])
        con.executemany("INSERT INTO detalle_entradas VALUES (?,?,?,?,?,?)", [
            (int(r["oc"]), int(_num(r["Línea"])), r["N° Entrada"], _iso(r["Fecha Entrada"]), _num(r["Cant. Ingresada"]),
             r["Almacén"]) for _, r in entradas.iterrows()])
        con.executemany("INSERT INTO detalle_facturas VALUES (?,?,?,?,?,?,?,?,?,?,?)", [
            (r["N° Interno FT"], int(r["oc"]), r["N° Factura"], r["Tipo"], r["Estado Factura"], _iso(r["Emisión FT"]),
             _iso(r["Contab. FT"]), _iso(r["Venc. FT"]), _iso(r["Fecha Pago"]), _num(r["Imp. Pendiente"]), r["RUC"])
            for _, r in facturas.iterrows()])
        filas = []
        for oc, g in lineas.groupby("oc"):
            c = g.iloc[0]
            datos = [(_num(x["Total Línea Sin IGV"]), *est_lin[(oc, x["Línea"])][:2], est_lin[(oc, x["Línea"])][2])
                     for _, x in g.iterrows()]
            est, avance = estado_oc(c["Tipo OC"], c["Estado OC"], datos)
            ing = sum(_num(x["Total Ingresado Línea"]) for _, x in g.iterrows())
            fechas = [fecha_ent.get((oc, x), "") for x in g["Línea"]]
            # proyecto / unidad / área de la línea de mayor monto
            mayor = g.assign(t=g["Total Línea Sin IGV"].map(_num)).sort_values("t", ascending=False).iloc[0]
            filas.append((int(oc), c["Tipo OC"], est,
                          None if c["Tipo OC"] == SERVICIO else int(ing > 0),
                          max(fechas) if any(fechas) else "", None if avance is None else round(avance, 4),
                          mayor["Proyecto"], mayor["Unidad Negocio"], mayor["Área"], len(g)))
        con.executemany("INSERT INTO detalle_oc VALUES (?,?,?,?,?,?,?,?,?,?)", filas)
        con.execute("INSERT INTO detalle_cargas(archivo, fecha_archivo, fecha_carga, filas, ocs) VALUES (?,?,?,?,?)",
                    (archivo, datetime.fromtimestamp(os.path.getmtime(archivo)).strftime("%Y-%m-%d %H:%M:%S")
                     if os.path.exists(archivo) else "", datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                     len(df), len(filas)))
        enlazar(con)
    return {"ocs": len(filas), "lineas": len(lineas), "entradas": len(entradas), "facturas": len(facturas)}


def _asegurar_columnas(con):
    """Bases creadas antes del estado por línea: agrega las columnas que faltan."""
    tiene = {r[1] for r in con.execute("PRAGMA table_info(detalle_lineas)")}
    for col, tipo in (("facturado", "REAL"), ("estado_linea", "TEXT")):
        if col not in tiene:
            con.execute(f"ALTER TABLE detalle_lineas ADD COLUMN {col} {tipo}")


def lineas_oc(con, oc):
    """Líneas de una OC con su estado de atención (para la consulta)."""
    _asegurar_columnas(con)
    return [dict(r) for r in con.execute("SELECT * FROM detalle_lineas WHERE numero_oc=? ORDER BY linea", (int(oc),))]


def enlazar(con):
    """Copia a sap_oc (último Control) el tipo de OC, el estado logístico y si tiene material ingresado."""
    con.execute("""UPDATE sap_oc SET tipo_oc = d.tipo_oc, estado_logistico = d.estado_logistico,
                   ingresado = d.ingresado, fecha_entrada = d.fecha_entrada, pct_atendido = d.pct_ingresado
                   FROM detalle_oc d WHERE d.numero_oc = sap_oc.numero_oc""")


def ultima_carga(con):
    return con.execute("SELECT * FROM detalle_cargas ORDER BY id DESC LIMIT 1").fetchone()


def ingresos(con):
    """{N° OC: 1/0/None} — material ingresado (None = servicio). Para la regla de nota de entrada."""
    return {r[0]: r[1] for r in con.execute("SELECT numero_oc, ingresado FROM detalle_oc")}
